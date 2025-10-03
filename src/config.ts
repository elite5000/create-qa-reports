import { config as loadEnvFile } from 'dotenv';

export interface Config {
  orgUrl: string;
  project: string;
  team: string;
  pat: string;
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

  const orgUrl = process.env.AZDO_ORG_URL;
  const project = process.env.AZDO_PROJECT;
  const team = process.env.AZDO_TEAM;
  const pat = process.env.AZDO_PAT;

  if (!orgUrl) {
    throw new Error('Missing env var AZDO_ORG_URL (e.g. https://dev.azure.com/your-org)');
  }
  if (!project) {
    throw new Error('Missing env var AZDO_PROJECT');
  }
  if (!team) {
    throw new Error('Missing env var AZDO_TEAM');
  }
  if (!pat) {
    throw new Error('Missing env var AZDO_PAT');
  }

  return { orgUrl, project, team, pat };
}