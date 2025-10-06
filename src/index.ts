import * as azdev from 'azure-devops-node-api';
import { IWorkItemTrackingApi } from 'azure-devops-node-api/WorkItemTrackingApi';
import { TeamContext } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { loadConfig } from './config';
import { parseCliOptions } from './cli/options';
import {
  fetchIterations,
  IterationWindow,
  selectIterations,
  IterationSelection,
} from './devops/iterations';
import {
  ConfluencePageResult,
  ConfluenceTemplate,
  fetchAndCacheConfluenceTemplate,
  syncReportToConfluence,
} from './integrations/confluence';
import { TestSuiteFinder } from './devops/testSuiteFinder';
import {
  ReportRow,
  TemplateRow,
  buildReportDocument,
  renderTable,
  writeReportDocument,
} from './reporting/htmlReport';
import {
  DEFAULT_WORK_ITEM_FIELDS,
  fetchWorkItemsForRange,
  extractAssignedTo,
} from './devops/workItems';

async function gatherReportRows(
  iterations: IterationWindow[],
  project: string,
  witApi: IWorkItemTrackingApi,
  suiteFinder: TestSuiteFinder
): Promise<ReportRow[]> {
  const seenIds = new Set<number>();
  const rows: ReportRow[] = [];

  for (const range of iterations) {
    const workItems = await fetchWorkItemsForRange(
      witApi,
      project,
      range,
      DEFAULT_WORK_ITEM_FIELDS
    );
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

  return rows;
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

  const iterationSelection: IterationSelection = selectIterations(
    allIterations,
    cliOptions.sprint
  );

  if (!iterationSelection.iterations.length) {
    console.log('Unable to determine a sprint to report on.');
    return;
  }

  console.log(`Generating QA report for sprint: ${iterationSelection.sprintLabel}`);

  const confluenceTemplate: ConfluenceTemplate = await fetchAndCacheConfluenceTemplate(config);

  const suiteFinder = new TestSuiteFinder(testPlanApi, config.project, config.orgUrl);
  const rows = await gatherReportRows(
    iterationSelection.iterations,
    config.project,
    witApi,
    suiteFinder
  );

  const templateRows: TemplateRow[] = rows.map((row) => ({
    id: row.id,
    assigned_to: row.assignedTo,
    test_plan_link: row.testSuiteLink ?? undefined,
  }));

  const reportHtml = buildReportDocument(
    confluenceTemplate.content,
    iterationSelection.sprintLabel,
    templateRows
  );
  const reportPaths = await writeReportDocument(iterationSelection.sprintLabel, reportHtml);
  const confluencePage: ConfluencePageResult = await syncReportToConfluence(
    config,
    iterationSelection.sprintLabel,
    reportHtml,
    confluenceTemplate
  );

  renderTable(rows);
  console.log(`Report ready at ${reportPaths.relative}`);
  console.log(`Confluence page URL: ${confluencePage.url}`);
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
