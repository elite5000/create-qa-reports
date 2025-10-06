import { formatTemplateValue } from '../utils/format';
import { TemplateRow } from './types';

const NO_DATA_MESSAGE = 'No completed Bugs or PBIs found for the configured sprints.';

function toHeaderTitle(key: string): string {
  const spaced = key.replace(/_/g, ' ');
  return spaced.replace(/\b([a-z])/g, (_, char: string) => char.toUpperCase());
}

export function renderTable(rows: TemplateRow[]): void {
  if (!rows.length) {
    console.log(NO_DATA_MESSAGE);
    return;
  }

  const columnKeys = Array.from(
    rows.reduce<Set<string>>((set, row) => {
      Object.keys(row).forEach((key) => set.add(key));
      return set;
    }, new Set<string>())
  );

  if (!columnKeys.length) {
    console.log(NO_DATA_MESSAGE);
    return;
  }

  const headers = columnKeys.map(toHeaderTitle);
  const data = rows.map((row) =>
    columnKeys.map((key) => formatTemplateValue(row[key]))
  );

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
