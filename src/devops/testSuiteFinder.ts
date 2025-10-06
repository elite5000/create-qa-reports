import { ITestPlanApi } from "azure-devops-node-api/TestPlanApi";
import { PagedList } from "azure-devops-node-api/interfaces/common/VSSInterfaces";
import * as TestPlanInterfaces from "azure-devops-node-api/interfaces/TestPlanInterfaces";

export class TestSuiteFinder {
    private readonly suiteTitleToLink = new Map<string, string>();
    private suitesLoaded = false;
    private readonly projectSegment: string;
    private readonly orgBaseUrl: string;

    constructor(
        private readonly api: ITestPlanApi,
        private readonly project: string,
        orgUrl: string
    ) {
        this.orgBaseUrl = orgUrl.replace(/\/+$/, "");
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
            continuationToken = (
                plansPage as PagedList<TestPlanInterfaces.TestPlan>
            ).continuationToken;
        } while (continuationToken);
    }

    private async indexSuitesForPlan(
        plan: TestPlanInterfaces.TestPlan
    ): Promise<void> {
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
            continuationToken = (
                suitesPage as PagedList<TestPlanInterfaces.TestSuite>
            ).continuationToken;
        } while (continuationToken);
    }

    private walkSuiteTree(
        plan: TestPlanInterfaces.TestPlan,
        suite: TestPlanInterfaces.TestSuite
    ): void {
        this.addSuite(plan, suite);
        for (const child of suite.children ?? []) {
            this.walkSuiteTree(plan, child);
        }
    }

    private addSuite(
        plan: TestPlanInterfaces.TestPlan,
        suite: TestPlanInterfaces.TestSuite
    ): void {
        const key = suite.name?.trim();
        if (!key || this.suiteTitleToLink.has(key)) {
            return;
        }

        const link = `${this.orgBaseUrl}/${this.projectSegment}/_testPlans/execute?planId=${plan.id}&suiteId=${suite.id}`;
        this.suiteTitleToLink.set(key, link);
    }
}
