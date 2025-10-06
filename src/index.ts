import * as azdev from "azure-devops-node-api";
import { IWorkItemTrackingApi } from "azure-devops-node-api/WorkItemTrackingApi";
import { TeamContext } from "azure-devops-node-api/interfaces/CoreInterfaces";
import { loadConfig } from "./config";
import { parseCliOptions } from "./cli/options";
import {
    fetchIterations,
    IterationWindow,
    selectIterations,
    IterationSelection,
} from "./devops/iterations";
import {
    DEFAULT_WORK_ITEM_FIELDS,
    fetchWorkItemsForRange,
    extractAssignedTo,
} from "./devops/workItems";
import { TestSuiteFinder } from "./devops/testSuiteFinder";
import {
    ConfluencePageResult,
    ConfluenceTemplate,
    fetchAndCacheConfluenceTemplate,
    syncReportToConfluence,
} from "./integrations/confluence";
import { buildReportDocument } from "./reporting/buildReportDocument";
import { renderTable } from "./reporting/renderTable";
import { writeReportDocument } from "./reporting/writeReportDocument";
import { TemplateRow } from "./reporting/types";

async function gatherTemplateRows(
    iterations: IterationWindow[],
    project: string,
    witApi: IWorkItemTrackingApi,
    suiteFinder: TestSuiteFinder
): Promise<TemplateRow[]> {
    const seenIds = new Set<number>();
    const rows: TemplateRow[] = [];

    for (const range of iterations) {
        const workItems = await fetchWorkItemsForRange(
            witApi,
            project,
            range,
            DEFAULT_WORK_ITEM_FIELDS
        );
        for (const item of workItems) {
            if (typeof item.id !== "number" || seenIds.has(item.id)) {
                continue;
            }
            seenIds.add(item.id);
            const assignedTo = extractAssignedTo(item);
            const testSuiteLink = await suiteFinder.findSuiteLinkByTitle(
                String(item.id)
            );
            rows.push({
                id: item.id,
                assigned_to: assignedTo,
                test_plan_link: testSuiteLink ?? undefined,
            });
        }
    }

    rows.sort((a, b) => {
        const assignedA = String(a.assigned_to ?? "");
        const assignedB = String(b.assigned_to ?? "");
        const assignedCompare = assignedA.localeCompare(assignedB);
        if (assignedCompare !== 0) {
            return assignedCompare;
        }
        const idA = typeof a.id === "number" ? a.id : Number(a.id ?? 0);
        const idB = typeof b.id === "number" ? b.id : Number(b.id ?? 0);
        return idA - idB;
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
        console.log(
            "No iterations with start and finish dates found for the configured team."
        );
        return;
    }

    const iterationSelection: IterationSelection = selectIterations(
        allIterations,
        cliOptions.sprint
    );

    if (!iterationSelection.iterations.length) {
        console.log("Unable to determine a sprint to report on.");
        return;
    }

    console.log(
        `Generating QA report for sprint: ${iterationSelection.sprintLabel}`
    );

    const confluenceTemplate: ConfluenceTemplate =
        await fetchAndCacheConfluenceTemplate(config);

    const suiteFinder = new TestSuiteFinder(
        testPlanApi,
        config.project,
        config.orgUrl
    );
    const templateRows = await gatherTemplateRows(
        iterationSelection.iterations,
        config.project,
        witApi,
        suiteFinder
    );

    renderTable(templateRows);

    const tableTitle = "Completed Bugs and PBIs";
    const reportHtml = buildReportDocument(
        confluenceTemplate.content,
        iterationSelection.sprintLabel,
        templateRows,
        { tableTitle }
    );
    const reportPaths = await writeReportDocument(
        iterationSelection.sprintLabel,
        reportHtml
    );
    const confluencePage: ConfluencePageResult = await syncReportToConfluence(
        config,
        iterationSelection.sprintLabel,
        reportHtml,
        confluenceTemplate
    );

    console.log(`Report ready at ${reportPaths.relative}`);
    console.log(`Confluence page URL: ${confluencePage.url}`);
}

main().catch((error) => {
    console.error("Failed to generate QA report:");
    if (error instanceof Error) {
        console.error(error.message);
    } else {
        console.error(error);
    }
    process.exitCode = 1;
});
