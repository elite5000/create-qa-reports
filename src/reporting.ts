import * as path from 'path';
import { writeFileIfChanged } from './fs-utils';

export interface ReportRow {
  id: number;
  assignedTo: string;
  testSuiteLink: string | null;
}

const REPORTS_DIR_RELATIVE = path.join('reports');
const REPORT_FILE_PREFIX = 'qa-report';

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

export function buildReportDocument(
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

export async function writeReportDocument(
  sprintLabel: string,
  html: string
): Promise<{ relative: string; absolute: string }> {
  const paths = computeReportPaths(sprintLabel);
  await writeFileIfChanged(paths.absolute, html);
  return paths;
}

export function renderTable(rows: ReportRow[]): void {
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
