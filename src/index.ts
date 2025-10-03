import * as azdev from 'azure-devops-node-api';
import { ITestPlanApi } from 'azure-devops-node-api/TestPlanApi';
import { IWorkApi } from 'azure-devops-node-api/WorkApi';
import { IWorkItemTrackingApi } from 'azure-devops-node-api/WorkItemTrackingApi';
import { TeamContext } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import {
  WorkItem,
  WorkItemErrorPolicy,
  WorkItemExpand,
  WorkItemReference,
} from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';
import { PagedList } from 'azure-devops-node-api/interfaces/common/VSSInterfaces';
import * as TestPlanInterfaces from 'azure-devops-node-api/interfaces/TestPlanInterfaces';
import { loadConfig } from './config';

interface ReportRow {
  id: number;
  assignedTo: string;
  testSuiteLink: string | null;
}

interface DateRange {
  start: Date;
  finish: Date;
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

async function fetchIterations(workApi: IWorkApi, teamContext: TeamContext): Promise<DateRange[]> {
  const timeframes: Array<'past' | 'current' | 'future'> = ['past', 'current', 'future'];
  const seen = new Set<string>();
  const ranges: DateRange[] = [];
  for (const timeframe of timeframes) {
    const iterations = await workApi.getTeamIterations(teamContext, timeframe);
    for (const iteration of iterations) {
      const iterationId = iteration.id ?? `${iteration.name}-${timeframe}`;
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
      if (!Number.isNaN(start.getTime()) && !Number.isNaN(finish.getTime())) {
        ranges.push({ start, finish });
      }
    }
  }
  return ranges;
}

async function fetchWorkItemsForRange(
  witApi: IWorkItemTrackingApi,
  project: string,
  range: DateRange
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
      WorkItemExpand.Fields,
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

  const ranges = await fetchIterations(workApi, teamContext);
  if (!ranges.length) {
    console.log('No iterations with start and finish dates found for the configured team.');
    return;
  }

  const suiteFinder = new TestSuiteFinder(testPlanApi, config.project, config.orgUrl);
  const seenIds = new Set<number>();
  const rows: ReportRow[] = [];

  for (const range of ranges) {
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

  renderTable(rows);
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
