const REGEX_SPECIAL_CHARACTERS = /[.*+?^${}()|[\]\\]/g;

const htmlEscapeMap: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;',
};

export interface TableSection {
  start: number;
  end: number;
  html: string;
}

export function escapeHtml(value: string): string {
  return String(value).replace(/[&<>"']/g, (char) => htmlEscapeMap[char] ?? char);
}

export function escapeRegExp(value: string): string {
  return value.replace(REGEX_SPECIAL_CHARACTERS, '\\$&');
}

export function findTemplateRow(html: string): string {
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

export function extractPlaceholders(templateRow: string): string[] {
  const matches = templateRow.match(/\{([^}]+)\}/g) ?? [];
  const placeholders = matches.map((placeholder) => placeholder.slice(1, -1));
  return Array.from(new Set(placeholders));
}

function findTableBounds(html: string, tableStart: number): TableSection {
  const tagRegex = /<\/?table\b[^>]*>/gi;
  tagRegex.lastIndex = tableStart;
  let depth = 0;
  let match: RegExpExecArray | null;
  while ((match = tagRegex.exec(html))) {
    const isOpening = match[0][1] !== '/';
    if (isOpening) {
      depth++;
    } else {
      depth--;
      if (depth === 0) {
        const end = match.index + match[0].length;
        return {
          start: tableStart,
          end,
          html: html.slice(tableStart, end),
        };
      }
    }
  }
  throw new Error('Failed to locate closing </table> tag.');
}

export function findTableAfterHeading(html: string, heading: string): TableSection | undefined {
  const normalizedHeading = heading.trim().toLowerCase();
  if (!normalizedHeading) {
    return undefined;
  }

  const headingRegex = /<h[1-6][^>]*>[\s\S]*?<\/h[1-6]>/gi;
  let match: RegExpExecArray | null;
  while ((match = headingRegex.exec(html))) {
    const textContent = match[0].replace(/<[^>]+>/g, '').trim().toLowerCase();
    if (textContent === normalizedHeading) {
      const tableStart = html.indexOf('<table', match.index + match[0].length);
      if (tableStart === -1) {
        return undefined;
      }
      return findTableBounds(html, tableStart);
    }
  }
  return undefined;
}

export function findFirstTable(html: string): TableSection {
  const tableStart = html.indexOf('<table');
  if (tableStart === -1) {
    throw new Error('Template does not contain any <table> elements.');
  }
  return findTableBounds(html, tableStart);
}
