import { TemplateValue } from '../utils/format';

export type TemplateRow = Record<string, TemplateValue>;

export interface BuildReportOptions {
  tableTitle?: string;
}
