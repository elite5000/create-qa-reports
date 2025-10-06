import { IWorkApi } from 'azure-devops-node-api/WorkApi';
import { TeamContext } from 'azure-devops-node-api/interfaces/CoreInterfaces';
import { TeamSettingsIteration } from 'azure-devops-node-api/interfaces/WorkInterfaces';

export interface IterationWindow {
  start: Date;
  finish: Date;
  name: string;
  path?: string;
  id?: string;
}

export interface IterationSelection {
  iterations: IterationWindow[];
  sprintLabel: string;
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

export function selectIterations(
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

export async function fetchIterations(
  workApi: IWorkApi,
  teamContext: TeamContext
): Promise<IterationWindow[]> {
  const seen = new Set<string>();
  const ranges: IterationWindow[] = [];

  const appendIterations = (
    iterations: TeamSettingsIteration[] | undefined,
    timeframeLabel?: string
  ) => {
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
