import { IWorkItemTrackingApi } from 'azure-devops-node-api/WorkItemTrackingApi';
import {
  WorkItem,
  WorkItemErrorPolicy,
  WorkItemExpand,
  WorkItemReference,
} from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';
import { IterationWindow } from './iterations';

export type WorkItemFieldMap = Record<string, string>;

export const DEFAULT_WORK_ITEM_FIELDS: WorkItemFieldMap = {
  id: 'System.Id',
  assigned_to: 'System.AssignedTo',
};

export async function fetchWorkItemsForRange(
  witApi: IWorkItemTrackingApi,
  project: string,
  range: IterationWindow,
  fieldMap: WorkItemFieldMap
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

  const requestedFields = Object.values(fieldMap);
  const fields = Array.from(new Set([...requestedFields, 'System.Id']));

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

export function extractAssignedTo(workItem: WorkItem): string {
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
