import * as path from 'path';
import { writeFileIfChanged } from '../utils/fs';
import { slugifyString } from '../utils/slugifyString';

export function computeReportFilePath(sprintLabel: string): {
  relative: string;
  absolute: string;
} {
  const fileName = `qa-report-${slugifyString(sprintLabel)}.html`;
  const relative = path.join('reports', fileName);
  const absolute = path.resolve(process.cwd(), relative);
  return { relative, absolute };
}

export async function writeReportDocument(
  sprintLabel: string,
  html: string
): Promise<{ relative: string; absolute: string }> {
  const paths = computeReportFilePath(sprintLabel);
  await writeFileIfChanged(paths.absolute, html);
  return paths;
}
