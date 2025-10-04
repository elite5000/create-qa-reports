import { promises as fs } from 'fs';
import * as path from 'path';
import * as azdev from 'azure-devops-node-api';
import { ITestPlanApi } from 'azure-devops-node-api/TestPlanApi';
import { IWorkApi } from 'azure-devops-node-api/WorkApi';
import { IWorkItemTrackingApi } from 'azure-devops-node-api/WorkItemTrackingApi';
import { TeamContext } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { TeamSettingsIteration } from 'azure-devops-node-api/interfaces/WorkInterfaces';
import {
  WorkItem,
  WorkItemErrorPolicy,
  WorkItemExpand,
  WorkItemReference,
} from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';
import { PagedList } from 'azure-devops-node-api/interfaces/common/VSSInterfaces';
import * as TestPlanInterfaces from 'azure-devops-node-api/interfaces/TestPlanInterfaces';
import { Config, loadConfig } from './config';

interface ReportRow {
  id: number;
  assignedTo: string;
  testSuiteLink: string | null;
}

interface IterationWindow {
  start: Date;
  finish: Date;
  name: string;
  path?: string;
  id?: string;
}

interface CliOptions {
  sprint?: string;
}

interface IterationSelection {
  iterations: IterationWindow[];
  sprintLabel: string;
}

interface ConfluenceContentResponse {
  id?: string;
  title?: string;
  body?: {
    storage?: {
      value?: string;
    };
  };
  version?: {
    number?: number;
  };
}

interface ConfluenceTemplate {
  content: string;
  version: number;
  title: string;
  cachedFilePath: string;
}

const TEMPLATE_CACHE_RELATIVE = path.join('templates', 'confluence-template.html');
const REPORTS_DIR_RELATIVE = path.join('reports');
const REPORT_FILE_PREFIX = 'qa-report';

interface ConfluencePageSummary {
  id: string;
  title: string;
  versionNumber: number;
}

interface ConfluenceSearchResponse {
  results?: Array<{
    id?: string;
    title?: string;
    version?: {
      number?: number;
    };
  }>;
}

function buildConfluenceBase(config: Config): { trimmed: string; withSlash: string } {
  const trimmed = config.confluenceBaseUrl.replace(/\/+$/, '');
  return { trimmed, withSlash: `${trimmed}/` };
}

function buildConfluenceAuthHeader(config: Config): string {
  return Buffer.from(
    `${config.confluenceEmail}:${config.confluenceApiToken}`,
    'utf8'
  ).toString('base64');
}

async function confluenceRequest(
  config: Config,
  relativePath: string,
  init: RequestInit = {}
): Promise<{ response: Response; text: string }> {
  const { withSlash } = buildConfluenceBase(config);
  const url = new URL(relativePath, withSlash);

  const headers = new Headers(init.headers ?? {});
  headers.set('Authorization', `Basic ${buildConfluenceAuthHeader(config)}`);
  if (!headers.has('Accept')) {
    headers.set('Accept', 'application/json; charset=utf-8');
  }
  if (init.body && !headers.has('Content-Type')) {
    headers.set('Content-Type', 'application/json; charset=utf-8');
  }

  const requestInit: RequestInit = {
    ...init,
    headers,
  };

  const response = await fetch(url.toString(), requestInit);
  const text = await response.text();

  if (!response.ok) {
    throw new Error(
      `Confluence request failed (${response.status} ${response.statusText}) for ${url.pathname}: ${text}`
    );
  }

  return { response, text };
}

function parseCliOptions(argv: string[]): CliOptions {
  const options: CliOptions = {};
  for (let i = 2; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--sprint' || arg === '-s') {
      if (i + 1 >= argv.length) {
        throw new Error('Expected a value after --sprint.');
      }
      options.sprint = argv[++i];
      continue;
    }
    if (arg.startsWith('--sprint=')) {
      options.sprint = arg.slice('--sprint='.length);
      continue;
    }
    if (arg === '--help' || arg === '-h') {
      console.log('Usage: npm run dev -- --sprint "Sprint Name"');
      process.exit(0);
    }
    console.warn(`Ignoring unknown argument: ${arg}`);
  }
  return options;
}

function normalizeIterationKey(value: string | undefined): string {
  return value?.trim().toLocaleLowerCase('en-US') ?? '';
}

function getCandidateKeys(iteration: IterationWindow): string[] {
  const keys = new Set<string>();
  keys.add(normalizeIterationKey(iteration.name));
  if (iteration.path) {
    keys.add(normalizeIterationKey(iteration.path));
    const segments = iteration.path.split('\\');
    for (const segment of segments) {
      keys.add(normalizeIterationKey(segment));
    }
  }
  if (iteration.id) {
    keys.add(normalizeIterationKey(iteration.id));
  }
  return Array.from(keys).filter((key) => key.length > 0);
}

function selectIterations(
  allIterations: IterationWindow[],
  sprintName?: string
): IterationSelection {
  if (!allIterations.length) {
    return { iterations: [], sprintLabel: sprintName ?? '' };
  }

  if (!sprintName) {
    const now = Date.now();
    const current = allIterations.find(
      (iteration) =>
        iteration.start.getTime() <= now && now <= iteration.finish.getTime()
    );
    if (current) {
      return { iterations: [current], sprintLabel: current.name };
    }
    const latest = allIterations[allIterations.length - 1];
    return { iterations: [latest], sprintLabel: latest.name };
  }

  const query = normalizeIterationKey(sprintName);
  const directMatch = allIterations.find((iteration) =>
    getCandidateKeys(iteration).some((key) => key === query)
  );
  if (directMatch) {
    return { iterations: [directMatch], sprintLabel: directMatch.name };
  }

  const partialMatch = allIterations.find((iteration) =>
    getCandidateKeys(iteration).some((key) => key.includes(query))
  );
  if (partialMatch) {
    return { iterations: [partialMatch], sprintLabel: partialMatch.name };
  }

  throw new Error(`Unable to find an iteration matching sprint "${sprintName}".`);
}

async function writeFileIfChanged(filePath: string, newContent: string): Promise<void> {
  let existingContent: string | undefined;
  try {
    existingContent = await fs.readFile(filePath, 'utf8');
  } catch (error) {
    const nodeError = error as NodeJS.ErrnoException;
    if (nodeError.code && nodeError.code !== 'ENOENT') {
      throw error;
    }
  }

  if (existingContent === newContent) {
    return;
  }

  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, newContent, 'utf8');
}

const htmlEscapeMap: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;',
};

function escapeHtml(value: string): string {
  return String(value).replace(/[&<>"']/g, (char) => htmlEscapeMap[char] ?? char);
}

function findTemplateRow(html: string): string {
  const rowRegex = /<tr\b[\s\S]*?<\/tr>/gi;
  const matches = html.match(rowRegex);
  if (!matches) {
    throw new Error('Template does not contain any table rows.');
  }

  const target = matches.find(
    (row) =>
      row.includes('{id}') && row.includes('{assigned_to}') && row.includes('{test_plan_link}')
  );

  if (!target) {
    throw new Error(
      'Unable to find template table row containing {id}, {assigned_to}, and {test_plan_link}. '
    );
  }

  return target;
}

function renderRowFromTemplate(templateRow: string, row: ReportRow): string {
  const idValue = escapeHtml(row.id.toString());
  const assignedValue = escapeHtml(row.assignedTo);
  const linkValue = row.testSuiteLink ? escapeHtml(row.testSuiteLink) : '';
  const linkDisplay = row.testSuiteLink ? linkValue : escapeHtml('Not Found');

  let rendered = templateRow;
  rendered = rendered.replace(/href\s*=\s*"\{test_plan_link\}"/gi, () =>
    row.testSuiteLink ? `href="${linkValue}"` : ''
  );
  rendered = rendered.replace(/\{id\}/g, idValue);
  rendered = rendered.replace(/\{assigned_to\}/g, assignedValue);
  rendered = rendered.replace(/\{test_plan_link\}/g, linkDisplay);
  return rendered;
}

function renderEmptyRow(templateRow: string): string {
  let rendered = templateRow;
  rendered = rendered.replace(/href\s*=\s*"\{test_plan_link\}"/gi, '');
  rendered = rendered.replace(/\{id\}/g, 'N/A');
  rendered = rendered.replace(/\{assigned_to\}/g, 'N/A');
  rendered = rendered.replace(/\{test_plan_link\}/g, 'No completed work items');
  return rendered;
}

function buildReportDocument(
  templateHtml: string,
  sprintLabel: string,
  rows: ReportRow[]
): string {
  let html = templateHtml.replace(/\{sprint\}/g, escapeHtml(sprintLabel));
  const templateRow = findTemplateRow(html);
  const tableRows = rows.length
    ? rows.map((row) => renderRowFromTemplate(templateRow, row)).join('')
    : renderEmptyRow(templateRow);

  html = html.replace(templateRow, tableRows);
  return html;
}

function slugifySprintLabel(sprintLabel: string): string {
  const trimmed = sprintLabel.trim();
  const withoutIllegal = trimmed.replace(/[<>:"/\\|?*]/g, '');
  const collapsedWhitespace = withoutIllegal.replace(/\s+/g, ' ').trim();
  const dashed = collapsedWhitespace.replace(/\s+/g, '-');
  const cleaned = dashed.replace(/-+/g, '-');
  return cleaned.toLowerCase() || 'report';
}

function computeReportPaths(sprintLabel: string): { relative: string; absolute: string } {
  const fileName = `${REPORT_FILE_PREFIX}-${slugifySprintLabel(sprintLabel)}.html`;
  const relative = path.join(REPORTS_DIR_RELATIVE, fileName);
  const absolute = path.resolve(process.cwd(), relative);
  return { relative, absolute };
}

async function writeReportDocument(
  sprintLabel: string,
  html: string
): Promise<{ relative: string; absolute: string }> {
  const paths = computeReportPaths(sprintLabel);
  await writeFileIfChanged(paths.absolute, html);
  return paths;
}

async function fetchAndCacheConfluenceTemplate(config: Config): Promise<ConfluenceTemplate> {
  const pageId = encodeURIComponent(config.confluencePageId);
  const { text } = await confluenceRequest(
    config,
    `rest/api/content/${pageId}?expand=body.storage,version`
  );

  let payload: ConfluenceContentResponse;
  try {
    payload = JSON.parse(text) as ConfluenceContentResponse;
  } catch (error) {
    throw new Error('Failed to parse Confluence response as JSON.');
  }

  const content = payload.body?.storage?.value;
  if (typeof content !== 'string') {
    throw new Error('Confluence template response did not include storage HTML content.');
  }

  const version = payload.version?.number ?? 0;
  const title = payload.title ?? 'Confluence Template';

  const cachePath = path.resolve(process.cwd(), TEMPLATE_CACHE_RELATIVE);
  await writeFileIfChanged(cachePath, content);

  const relativePath = path.relative(process.cwd(), cachePath) || cachePath;
  console.log(
    `Fetched Confluence template "${title}" (version ${version}). Cached at ${relativePath}.`
  );

  return { content, version, title, cachedFilePath: cachePath };
}

function computeReportTitle(templateTitle: string, sprintLabel: string): string {
  const baseTitle = templateTitle?.trim() || 'QA Report';
  const sprint = sprintLabel.trim();
  return sprint ? `${baseTitle} - ${sprint}` : baseTitle;
}

async function findExistingReportPage(
  config: Config,
  title: string
): Promise<ConfluencePageSummary | null> {
  const spaceKey = encodeURIComponent(config.confluenceSpaceKey);
  const encodedTitle = encodeURIComponent(title);
  const { text } = await confluenceRequest(
    config,
    `rest/api/content?spaceKey=${spaceKey}&title=${encodedTitle}&expand=version`
  );

  let payload: ConfluenceSearchResponse;
  try {
    payload = JSON.parse(text) as ConfluenceSearchResponse;
  } catch (error) {
    throw new Error('Failed to parse Confluence search response as JSON.');
  }

  const match = payload.results?.find((result) => {
    if (!result?.title || !result.id) {
      return false;
    }
    return result.title.trim().toLocaleLowerCase('en-US') === title.trim().toLocaleLowerCase('en-US');
  });

  if (!match?.id) {
    return null;
  }

  return {
    id: match.id,
    title: match.title ?? title,
    versionNumber: match.version?.number ?? 1,
  };
}

async function createConfluenceReportPage(
  config: Config,
  title: string,
  html: string
): Promise<ConfluencePageSummary> {
  const payload: Record<string, unknown> = {
    type: 'page',
    title,
    space: { key: config.confluenceSpaceKey },
    body: {
      storage: {
        value: html,
        representation: 'storage',
      },
    },
  };

  if (config.confluenceParentPageId) {
    payload.ancestors = [{ id: config.confluenceParentPageId }];
  }

  const { text } = await confluenceRequest(config, 'rest/api/content', {
    method: 'POST',
    body: JSON.stringify(payload),
  });

  let response: ConfluenceContentResponse;
  try {
    response = JSON.parse(text) as ConfluenceContentResponse;
  } catch (error) {
    throw new Error('Failed to parse Confluence create response as JSON.');
  }

  if (!response.id) {
    throw new Error('Confluence did not return an ID for the created page.');
  }

  const versionNumber = response.version?.number ?? 1;
  return {
    id: response.id,
    title: response.title ?? title,
    versionNumber,
  };
}

async function updateConfluenceReportPage(
  config: Config,
  existing: ConfluencePageSummary,
  title: string,
  html: string
): Promise<ConfluencePageSummary> {
  const nextVersion = existing.versionNumber + 1;
  const payload: Record<string, unknown> = {
    id: existing.id,
    type: 'page',
    title,
    version: {
      number: nextVersion,
    },
    body: {
      storage: {
        value: html,
        representation: 'storage',
      },
    },
  };

  if (config.confluenceParentPageId) {
    payload.ancestors = [{ id: config.confluenceParentPageId }];
  }

  const { text } = await confluenceRequest(
    config,
    `rest/api/content/${encodeURIComponent(existing.id)}`,
    {
      method: 'PUT',
      body: JSON.stringify(payload),
    }
  );

  let response: ConfluenceContentResponse;
  try {
    response = JSON.parse(text) as ConfluenceContentResponse;
  } catch (error) {
    throw new Error('Failed to parse Confluence update response as JSON.');
  }

  const versionNumber = response.version?.number ?? nextVersion;
  return {
    id: existing.id,
    title: response.title ?? title,
    versionNumber,
  };
}

async function syncReportToConfluence(
  config: Config,
  sprintLabel: string,
  html: string,
  template: ConfluenceTemplate
): Promise<ConfluencePageSummary> {
  const title = computeReportTitle(template.title, sprintLabel);
  const existing = await findExistingReportPage(config, title);
  if (!existing) {
    const created = await createConfluenceReportPage(config, title, html);
    console.log(`Created Confluence page "${created.title}" (ID ${created.id}).`);
    return created;
  }

  const updated = await updateConfluenceReportPage(config, existing, title, html);
  console.log(
    `Updated Confluence page "${updated.title}" (ID ${updated.id}) to version ${updated.versionNumber}.`
  );
  return updated;
}

class TestSuiteFinder {
  private readonly suiteTitleToLink = new Map<string, string>();
  private suitesLoaded = false;
  private readonly projectSegment: string;
  private readonly orgBaseUrl: string;

  constructor(
    private readonly api: ITestPlanApi,
    private readonly project: string,
    orgUrl: string
  ) {
    this.orgBaseUrl = orgUrl.replace(/\/+$/, '');
    this.projectSegment = encodeURIComponent(project);
  }

  async findSuiteLinkByTitle(title: string): Promise<string | null> {
    if (!title) {
      return null;
    }

    if (!this.suitesLoaded) {
      await this.loadAllSuites();
      this.suitesLoaded = true;
    }

    const trimmed = title.trim();
    return this.suiteTitleToLink.get(trimmed) ?? null;
  }

  private async loadAllSuites(): Promise<void> {
    let continuationToken: string | undefined;
    do {
      const plansPage = await this.api.getTestPlans(
        this.project,
        undefined,
        continuationToken,
        false,
        false
      );
      const plans = plansPage ?? [];
      for (const plan of plans) {
        await this.indexSuitesForPlan(plan);
      }
      continuationToken = (plansPage as PagedList<TestPlanInterfaces.TestPlan>).continuationToken;
    } while (continuationToken);
  }

  private async indexSuitesForPlan(plan: TestPlanInterfaces.TestPlan): Promise<void> {
    let continuationToken: string | undefined;
    do {
      const suitesPage = await this.api.getTestSuitesForPlan(
        this.project,
        plan.id,
        TestPlanInterfaces.SuiteExpand.Children,
        continuationToken,
        true
      );
      const suites = suitesPage ?? [];
      for (const suite of suites) {
        this.walkSuiteTree(plan, suite);
      }
      continuationToken = (suitesPage as PagedList<TestPlanInterfaces.TestSuite>).continuationToken;
    } while (continuationToken);
  }

  private walkSuiteTree(plan: TestPlanInterfaces.TestPlan, suite: TestPlanInterfaces.TestSuite): void {
    this.addSuite(plan, suite);
    for (const child of suite.children ?? []) {
      this.walkSuiteTree(plan, child);
    }
  }

  private addSuite(plan: TestPlanInterfaces.TestPlan, suite: TestPlanInterfaces.TestSuite): void {
    const key = suite.name?.trim();
    if (!key || this.suiteTitleToLink.has(key)) {
      return;
    }

    const link = `${this.orgBaseUrl}/${this.projectSegment}/_testPlans/execute?planId=${plan.id}&suiteId=${suite.id}`;
    this.suiteTitleToLink.set(key, link);
  }
}

async function fetchIterations(
  workApi: IWorkApi,
  teamContext: TeamContext
): Promise<IterationWindow[]> {
  const seen = new Set<string>();
  const ranges: IterationWindow[] = [];

  const appendIterations = (iterations: TeamSettingsIteration[] | undefined, timeframeLabel?: string) => {
    if (!iterations?.length) {
      return;
    }

    for (const iteration of iterations) {
      const iterationId = iteration.id ?? `${iteration.name}-${timeframeLabel ?? 'all'}`;
      if (!iterationId || seen.has(iterationId)) {
        continue;
      }
      seen.add(iterationId);
      const attributes = iteration.attributes;
      if (!attributes?.startDate || !attributes?.finishDate) {
        continue;
      }
      const start = new Date(attributes.startDate);
      const finish = new Date(attributes.finishDate);
      const name = iteration.name?.trim();
      if (!name) {
        continue;
      }
      if (!Number.isNaN(start.getTime()) && !Number.isNaN(finish.getTime())) {
        ranges.push({
          id: iteration.id ?? undefined,
          name,
          path: iteration.path ?? undefined,
          start,
          finish,
        });
      }
    }
  };

  const timeframes: Array<'past' | 'current' | 'future'> = ['past', 'current', 'future'];
  for (const timeframe of timeframes) {
    try {
      const iterations = await workApi.getTeamIterations(teamContext, timeframe);
      appendIterations(iterations, timeframe);
    } catch (error) {
      if (!isUnsupportedTimeframeError(error)) {
        throw error;
      }
      const fallbackIterations = await workApi.getTeamIterations(teamContext);
      appendIterations(fallbackIterations);
      break;
    }
  }

  if (!ranges.length) {
    const allIterations = await workApi.getTeamIterations(teamContext);
    appendIterations(allIterations);
  }

  ranges.sort((a, b) => a.start.getTime() - b.start.getTime());
  return ranges;
}

function isUnsupportedTimeframeError(error: unknown): boolean {
  if (!error || typeof error !== 'object') {
    return false;
  }

  const message = (error as { message?: unknown }).message;
  if (typeof message !== 'string') {
    return false;
  }

  const normalized = message.toLowerCase();
  return normalized.includes('timeframe');
}

async function fetchWorkItemsForRange(
  witApi: IWorkItemTrackingApi,
  project: string,
  range: IterationWindow
): Promise<WorkItem[]> {
  const startIso = range.start.toISOString();
  const finishIso = range.finish.toISOString();
  const wiql = {
    query: `
      SELECT [System.Id]
      FROM WorkItems
      WHERE [System.TeamProject] = @project
        AND [System.WorkItemType] IN ('Bug', 'Product Backlog Item')
        AND [Microsoft.VSTS.Common.ClosedDate] >= '${startIso}'
        AND [Microsoft.VSTS.Common.ClosedDate] <= '${finishIso}'
    `,
  };

  const result = await witApi.queryByWiql(wiql, { project });
  const ids = (result.workItems ?? [])
    .map((ref: WorkItemReference) => ref.id)
    .filter((id): id is number => typeof id === 'number');

  if (!ids.length) {
    return [];
  }

  const fields = [
    'System.Id',
    'System.AssignedTo',
    'Microsoft.VSTS.Common.ClosedDate',
  ];

  const items: WorkItem[] = [];
  const chunkSize = 200;
  for (let i = 0; i < ids.length; i += chunkSize) {
    const chunk = ids.slice(i, i + chunkSize);
    const chunkItems = await witApi.getWorkItems(
      chunk,
      fields,
      undefined,
      WorkItemExpand.None,
      WorkItemErrorPolicy.Omit
    );
    items.push(...chunkItems);
  }

  return items;
}

function extractAssignedTo(workItem: WorkItem): string {
  const raw = workItem.fields?.['System.AssignedTo'];
  if (!raw) {
    return 'Unassigned';
  }
  if (typeof raw === 'string') {
    return raw;
  }
  const identity = raw as { displayName?: string; uniqueName?: string };
  return identity.displayName || identity.uniqueName || 'Unassigned';
}

function renderTable(rows: ReportRow[]): void {
  if (!rows.length) {
    console.log('No completed Bugs or PBIs found for the configured sprints.');
    return;
  }

  const headers = ['ID', 'Assigned To', 'Test Suite Link'];
  const data = rows.map((row) => [
    row.id.toString(),
    row.assignedTo,
    row.testSuiteLink ?? 'Not Found',
  ]);

  const colWidths = headers.map((header, idx) =>
    Math.max(
      header.length,
      ...data.map((line) => line[idx].length)
    )
  );

  const divider = `+${colWidths.map((w) => '-'.repeat(w + 2)).join('+')}+`;

  const lines: string[] = [];
  lines.push(divider);
  lines.push(
    `| ${headers.map((header, idx) => header.padEnd(colWidths[idx])).join(' | ')} |`
  );
  lines.push(divider);
  for (const row of data) {
    lines.push(
      `| ${row.map((cell, idx) => cell.padEnd(colWidths[idx])).join(' | ')} |`
    );
  }
  lines.push(divider);

  console.log(lines.join('\n'));
}

async function main(): Promise<void> {
  const cliOptions = parseCliOptions(process.argv);
  const config = loadConfig();

  const authHandler = azdev.getPersonalAccessTokenHandler(config.pat);
  const connection = new azdev.WebApi(config.orgUrl, authHandler);

  const teamContext: TeamContext = {
    project: config.project,
    team: config.team,
  };

  const [workApi, witApi, testPlanApi] = await Promise.all([
    connection.getWorkApi(),
    connection.getWorkItemTrackingApi(),
    connection.getTestPlanApi(),
  ]);

  const allIterations = await fetchIterations(workApi, teamContext);
  if (!allIterations.length) {
    console.log('No iterations with start and finish dates found for the configured team.');
    return;
  }

  const { iterations: selectedIterations, sprintLabel } = selectIterations(
    allIterations,
    cliOptions.sprint
  );

  if (!selectedIterations.length) {
    console.log('Unable to determine a sprint to report on.');
    return;
  }

  console.log(`Generating QA report for sprint: ${sprintLabel}`);

  const confluenceTemplate = await fetchAndCacheConfluenceTemplate(config);

  const suiteFinder = new TestSuiteFinder(testPlanApi, config.project, config.orgUrl);
  const seenIds = new Set<number>();
  const rows: ReportRow[] = [];

  for (const range of selectedIterations) {
    const workItems = await fetchWorkItemsForRange(witApi, config.project, range);
    for (const item of workItems) {
      if (typeof item.id !== 'number' || seenIds.has(item.id)) {
        continue;
      }
      seenIds.add(item.id);
      const assignedTo = extractAssignedTo(item);
      const testSuiteLink = await suiteFinder.findSuiteLinkByTitle(String(item.id));
      rows.push({
        id: item.id,
        assignedTo,
        testSuiteLink,
      });
    }
  }

  rows.sort((a, b) => {
    const assignedCompare = a.assignedTo.localeCompare(b.assignedTo);
    if (assignedCompare !== 0) {
      return assignedCompare;
    }
    return a.id - b.id;
  });

  const reportHtml = buildReportDocument(
    confluenceTemplate.content,
    sprintLabel,
    rows
  );
  const reportPaths = await writeReportDocument(sprintLabel, reportHtml);
  const confluencePage = await syncReportToConfluence(
    config,
    sprintLabel,
    reportHtml,
    confluenceTemplate
  );

  renderTable(rows);
  console.log(`Report ready at ${reportPaths.relative}`);
  console.log(`Confluence page URL: ${config.confluenceBaseUrl.replace(/\/+$/, '')}/spaces/${encodeURIComponent(config.confluenceSpaceKey)}/pages/${encodeURIComponent(confluencePage.id)}`);
}

main().catch((error) => {
  console.error('Failed to generate QA report:');
  if (error instanceof Error) {
    console.error(error.message);
  } else {
    console.error(error);
  }
  process.exitCode = 1;
});
