import { ReportRow } from './types';

const NO_DATA_MESSAGE = 'No completed Bugs or PBIs found for the configured sprints.';

export function renderTable(rows: ReportRow[]): void {
  if (!rows.length) {
    console.log(NO_DATA_MESSAGE);
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
