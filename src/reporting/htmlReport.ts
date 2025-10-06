import * as path from 'path';
import { writeFileIfChanged } from '../utils/fs';
import { formatTemplateValue, TemplateValue } from '../utils/format';

export interface ReportRow {
  id: number;
  assignedTo: string;
  testSuiteLink: string | null;
}

export type TemplateRow = Record<string, TemplateValue>;

const REPORTS_DIR_RELATIVE = path.join('reports');
const REPORT_FILE_PREFIX = 'qa-report';
const NO_WORK_ITEMS_MESSAGE = 'No work items found';

const htmlEscapeMap: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;',
};

const REGEX_SPECIAL_CHARACTERS = /[.*+?^${}()|[\]\\]/g;

function escapeHtml(value: string): string {
  return String(value).replace(/[&<>"']/g, (char) => htmlEscapeMap[char] ?? char);
}

function escapeRegExp(value: string): string {
  return value.replace(REGEX_SPECIAL_CHARACTERS, '\\$&');
}

function findTemplateRow(html: string): string {
  const rowRegex = /<tr\b[\s\S]*?<\/tr>/gi;
  const matches = html.match(rowRegex);
  if (!matches) {
    throw new Error('Template does not contain any table rows.');
  }

  const target = matches.find((row) => /\{[^}]+\}/.test(row));
  if (!target) {
    throw new Error('Unable to find template table row containing placeholders.');
  }

  return target;
}

function extractPlaceholders(templateRow: string): string[] {
  const matches = templateRow.match(/\{([^}]+)\}/g) ?? [];
  const placeholders = matches.map((placeholder) => placeholder.slice(1, -1));
  return Array.from(new Set(placeholders));
}

function replacePlaceholder(
  source: string,
  key: string,
  value: string,
  removeHrefWhenMissing: boolean
): string {
  const escapedKey = escapeRegExp(key);
  const hrefPattern = new RegExp(`href\\s*=\\s*"\\{${escapedKey}\\}"`, 'gi');
  const shouldRemoveHref =
    removeHrefWhenMissing && (value === NO_WORK_ITEMS_MESSAGE || value === 'N/A');
  source = source.replace(hrefPattern, shouldRemoveHref ? '' : `href="${value}"`);

  const placeholderPattern = new RegExp(`\\{${escapedKey}\\}`, 'g');
  return source.replace(placeholderPattern, value);
}

function renderRowFromTemplate(
  templateRow: string,
  placeholders: string[],
  row: TemplateRow
): string {
  let rendered = templateRow;
  for (const key of placeholders) {
    const formatted = formatTemplateValue(row[key]);
    const escaped = escapeHtml(formatted);
    rendered = replacePlaceholder(rendered, key, escaped, true);
  }
  return rendered;
}

function renderEmptyRow(templateRow: string, placeholders: string[]): string {
  let rendered = templateRow;
  for (const key of placeholders) {
    const escaped = escapeHtml(NO_WORK_ITEMS_MESSAGE);
    rendered = replacePlaceholder(rendered, key, escaped, true);
  }
  return rendered;
}

export function buildReportDocument(
  templateHtml: string,
  sprintLabel: string,
  rows: TemplateRow[]
): string {
  let html = templateHtml.replace(/\{sprint\}/g, escapeHtml(sprintLabel));
  const templateRow = findTemplateRow(html);
  const placeholders = extractPlaceholders(templateRow);
  const tableRows = rows.length
    ? rows
        .map((row) => renderRowFromTemplate(templateRow, placeholders, row))
        .join('')
    : renderEmptyRow(templateRow, placeholders);

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
    row.testSuiteLink ?? 'N/A',
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
