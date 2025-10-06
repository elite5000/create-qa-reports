import { config as loadEnvFile } from 'dotenv';

export interface Config {
  orgUrl: string;
  project: string;
  team: string;
  pat: string;
  confluenceBaseUrl: string;
  confluencePageId: string;
  confluenceEmail: string;
  confluenceApiToken: string;
  confluenceSpaceKey?: string;
  confluenceParentPageId?: string;
}

let envLoaded = false;

function ensureEnvLoaded(): void {
  if (envLoaded) {
    return;
  }

  const result = loadEnvFile({ quiet: true });
  if (result.error) {
    const err = result.error as NodeJS.ErrnoException;
    if (err.code !== 'ENOENT') {
      throw new Error(`Failed to load .env file: ${result.error.message}`);
    }
  }

  envLoaded = true;
}

export function loadConfig(): Config {
  ensureEnvLoaded();

  const requireEnv = (name: string, message?: string): string => {
    const value = process.env[name]?.trim();
    if (!value) {
      throw new Error(message ?? `Missing env var ${name}`);
    }
    return value;
  };

  const orgUrl = requireEnv(
    'AZDO_ORG_URL',
    'Missing env var AZDO_ORG_URL (e.g. https://dev.azure.com/your-org)'
  );
  const project = requireEnv('AZDO_PROJECT');
  const team = requireEnv('AZDO_TEAM');
  const pat = requireEnv('AZDO_PAT');
  const confluenceBaseUrl = requireEnv(
    'CONFLUENCE_BASE_URL',
    'Missing env var CONFLUENCE_BASE_URL (e.g. https://your-domain.atlassian.net/wiki)'
  );
  const confluencePageId = requireEnv('CONFLUENCE_PAGE_ID');
  const confluenceEmail = requireEnv('CONFLUENCE_EMAIL');
  const confluenceApiToken = requireEnv('CONFLUENCE_API_TOKEN');
  const confluenceSpaceKey = process.env.CONFLUENCE_SPACE_KEY?.trim();
  const confluenceParentPageId = process.env.CONFLUENCE_PARENT_PAGE_ID?.trim();

  return {
    orgUrl,
    project,
    team,
    pat,
    confluenceBaseUrl,
    confluencePageId,
    confluenceEmail,
    confluenceApiToken,
    confluenceSpaceKey: confluenceSpaceKey || undefined,
    confluenceParentPageId: confluenceParentPageId || undefined,
  };
}
