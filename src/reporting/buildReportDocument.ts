import { formatTemplateValue } from '../utils/format';
import { slugifyString } from '../utils/slugifyString';
import {
  extractPlaceholders,
  findFirstTable,
  findTableAfterHeading,
  findTemplateRow,
  TableSection,
  escapeHtml,
  escapeRegExp,
} from './findElements';
import { BuildReportOptions, TemplateRow } from './types';

const NO_WORK_ITEMS_MESSAGE = 'No work items found';

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
  let updated = source.replace(hrefPattern, shouldRemoveHref ? '' : `href="${value}"`);

  const placeholderPattern = new RegExp(`\\{${escapedKey}\\}`, 'g');
  updated = updated.replace(placeholderPattern, value);
  return updated;
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

function getTableSection(
  html: string,
  tableTitle: string | undefined
): { section: TableSection; headingExists: boolean } {
  if (tableTitle) {
    const section = findTableAfterHeading(html, tableTitle);
    if (section) {
      return { section, headingExists: true };
    }
  }
  return { section: findFirstTable(html), headingExists: false };
}

export function buildReportDocument(
  templateHtml: string,
  sprintLabel: string,
  rows: TemplateRow[],
  options: BuildReportOptions = {}
): string {
  let html = templateHtml.replace(/\{sprint\}/g, escapeHtml(sprintLabel));

  const tableTitle = options.tableTitle?.trim();
  const { section, headingExists } = getTableSection(html, tableTitle);

  const templateRow = findTemplateRow(section.html);
  const placeholders = extractPlaceholders(templateRow);
  const tableRows = rows.length
    ? rows
        .map((row) => renderRowFromTemplate(templateRow, placeholders, row))
        .join('')
    : renderEmptyRow(templateRow, placeholders);

  const updatedTableHtml = section.html.replace(templateRow, tableRows);
  html = html.slice(0, section.start) + updatedTableHtml + html.slice(section.end);

  if (tableTitle && !headingExists) {
    const sanitizedTitle = escapeHtml(tableTitle);
    const heading = `<h3 data-table-slug="${slugifyString(tableTitle)}">${sanitizedTitle}</h3>`;
    html = html.slice(0, section.start) + heading + html.slice(section.start);
  }

  return html;
}
