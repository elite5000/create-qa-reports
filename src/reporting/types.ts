import { TemplateValue } from '../utils/format';

export interface ReportRow {
  id: number;
  assignedTo: string;
  testSuiteLink: string | null;
}

export type TemplateRow = Record<string, TemplateValue>;

export interface BuildReportOptions {
  tableTitle?: string;
}
