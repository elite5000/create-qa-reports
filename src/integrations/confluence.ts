import * as path from "path";
import { Config } from "../config";
import { writeFileIfChanged } from "../utils/fs";

export interface ConfluenceTemplate {
    content: string;
    version: number;
    title: string;
    cachedFilePath: string;
    spaceKey: string;
    parentPageId?: string;
}

export interface ConfluencePageResult {
    id: string;
    title: string;
    versionNumber: number;
    url: string;
    spaceKey: string;
    parentPageId?: string;
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
    space?: {
        key?: string;
    };
    ancestors?: Array<{
        id?: string;
    }>;
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

interface ConfluencePageSummary {
    id: string;
    title: string;
    versionNumber: number;
}

const TEMPLATE_CACHE_RELATIVE = path.join(
    "templates",
    "confluence-template.html"
);

function buildConfluenceBase(config: Config): {
    trimmed: string;
    withSlash: string;
} {
    const trimmed = config.confluenceBaseUrl.replace(/\/+$/, "");
    return { trimmed, withSlash: `${trimmed}/` };
}

function buildConfluenceAuthHeader(config: Config): string {
    return Buffer.from(
        `${config.confluenceEmail}:${config.confluenceApiToken}`,
        "utf8"
    ).toString("base64");
}

async function confluenceRequest(
    config: Config,
    relativePath: string,
    init: RequestInit = {}
): Promise<{ response: Response; text: string; url: URL }> {
    const { withSlash } = buildConfluenceBase(config);
    const url = new URL(relativePath, withSlash);

    const headers = new Headers(init.headers ?? {});
    headers.set("Authorization", `Basic ${buildConfluenceAuthHeader(config)}`);
    if (!headers.has("Accept")) {
        headers.set("Accept", "application/json; charset=utf-8");
    }
    if (init.body && !headers.has("Content-Type")) {
        headers.set("Content-Type", "application/json; charset=utf-8");
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

    return { response, text, url };
}

export async function fetchAndCacheConfluenceTemplate(
    config: Config
): Promise<ConfluenceTemplate> {
    const pageId = encodeURIComponent(config.confluencePageId);
    const { text } = await confluenceRequest(
        config,
        `rest/api/content/${pageId}?expand=body.storage,version,space,ancestors`
    );

    let payload: ConfluenceContentResponse;
    try {
        payload = JSON.parse(text) as ConfluenceContentResponse;
    } catch (error) {
        throw new Error("Failed to parse Confluence response as JSON.");
    }

    const content = payload.body?.storage?.value;
    if (typeof content !== "string") {
        throw new Error(
            "Confluence template response did not include storage HTML content."
        );
    }

    const version = payload.version?.number ?? 0;
    const title = payload.title ?? "Confluence Template";
    const spaceKey = payload.space?.key ?? config.confluenceSpaceKey;
    if (!spaceKey) {
        throw new Error(
            "Unable to determine Confluence space key from template. Set CONFLUENCE_SPACE_KEY to override."
        );
    }

    const ancestors = payload.ancestors ?? [];
    const templateParentId = ancestors.length
        ? ancestors[ancestors.length - 1]?.id ?? undefined
        : undefined;

    const cachePath = path.resolve(process.cwd(), TEMPLATE_CACHE_RELATIVE);
    await writeFileIfChanged(cachePath, content);

    const relativePath = path.relative(process.cwd(), cachePath) || cachePath;
    console.log(
        `Fetched Confluence template "${title}" (version ${version}). Cached at ${relativePath}.`
    );

    return {
        content,
        version,
        title,
        cachedFilePath: cachePath,
        spaceKey,
        parentPageId: templateParentId,
    };
}

export function computeReportTitle(
    templateTitle: string,
    sprintLabel: string
): string {
    const baseTitle = templateTitle?.trim() || "QA Report";
    const sprint = sprintLabel.trim();
    return sprint ? baseTitle.replace("{sprint}", sprint) : baseTitle;
}

async function findExistingReportPage(
    config: Config,
    spaceKey: string,
    title: string
): Promise<ConfluencePageSummary | null> {
    const encodedSpaceKey = encodeURIComponent(spaceKey);
    const encodedTitle = encodeURIComponent(title);
    const { text } = await confluenceRequest(
        config,
        `rest/api/content?spaceKey=${encodedSpaceKey}&title=${encodedTitle}&expand=version`
    );

    let payload: ConfluenceSearchResponse;
    try {
        payload = JSON.parse(text) as ConfluenceSearchResponse;
    } catch (error) {
        throw new Error("Failed to parse Confluence search response as JSON.");
    }

    const match = payload.results?.find((result) => {
        if (!result?.title || !result.id) {
            return false;
        }
        return (
            result.title.trim().toLocaleLowerCase("en-US") ===
            title.trim().toLocaleLowerCase("en-US")
        );
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
    spaceKey: string,
    title: string,
    html: string,
    parentPageId?: string
): Promise<ConfluencePageSummary> {
    const payload: Record<string, unknown> = {
        type: "page",
        title,
        space: { key: spaceKey },
        body: {
            storage: {
                value: html,
                representation: "storage",
            },
        },
    };

    const parentId = config.confluenceParentPageId ?? parentPageId;
    if (parentId) {
        payload.ancestors = [{ id: parentId }];
    }

    const { text } = await confluenceRequest(config, "rest/api/content", {
        method: "POST",
        body: JSON.stringify(payload),
    });

    let response: ConfluenceContentResponse;
    try {
        response = JSON.parse(text) as ConfluenceContentResponse;
    } catch (error) {
        throw new Error("Failed to parse Confluence create response as JSON.");
    }

    if (!response.id) {
        throw new Error(
            "Confluence did not return an ID for the created page."
        );
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
    spaceKey: string,
    title: string,
    html: string,
    parentPageId?: string
): Promise<ConfluencePageSummary> {
    const nextVersion = existing.versionNumber + 1;
    const payload: Record<string, unknown> = {
        id: existing.id,
        type: "page",
        title,
        version: {
            number: nextVersion,
        },
        space: { key: spaceKey },
        body: {
            storage: {
                value: html,
                representation: "storage",
            },
        },
    };

    const parentId = config.confluenceParentPageId ?? parentPageId;
    if (parentId) {
        payload.ancestors = [{ id: parentId }];
    }

    const { text } = await confluenceRequest(
        config,
        `rest/api/content/${encodeURIComponent(existing.id)}`,
        {
            method: "PUT",
            body: JSON.stringify(payload),
        }
    );

    let response: ConfluenceContentResponse;
    try {
        response = JSON.parse(text) as ConfluenceContentResponse;
    } catch (error) {
        throw new Error("Failed to parse Confluence update response as JSON.");
    }

    const versionNumber = response.version?.number ?? nextVersion;
    return {
        id: existing.id,
        title: response.title ?? title,
        versionNumber,
    };
}

function buildPageUrl(
    config: Config,
    spaceKey: string,
    pageId: string
): string {
    const { trimmed } = buildConfluenceBase(config);
    const encodedId = encodeURIComponent(pageId);
    return `${trimmed}/spaces/${encodeURIComponent(
        spaceKey
    )}/pages/${encodedId}`;
}

export async function syncReportToConfluence(
    config: Config,
    sprintLabel: string,
    html: string,
    template: ConfluenceTemplate
): Promise<ConfluencePageResult> {
    const spaceKey = template.spaceKey;
    const parentPageId = config.confluenceParentPageId ?? template.parentPageId;
    const title = computeReportTitle(template.title, sprintLabel);
    const existing = await findExistingReportPage(config, spaceKey, title);
    const page = existing
        ? await updateConfluenceReportPage(
              config,
              existing,
              spaceKey,
              title,
              html,
              parentPageId
          )
        : await createConfluenceReportPage(
              config,
              spaceKey,
              title,
              html,
              parentPageId
          );

    if (existing) {
        console.log(
            `Updated Confluence page "${page.title}" (ID ${page.id}) to version ${page.versionNumber}.`
        );
    } else {
        console.log(`Created Confluence page "${page.title}" (ID ${page.id}).`);
    }

    return {
        ...page,
        spaceKey,
        parentPageId,
        url: buildPageUrl(config, spaceKey, page.id),
    };
}
