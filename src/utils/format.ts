export type TemplateValue = string | number | Date | null | undefined;

const MISSING_VALUE_TEXT = 'N/A';

export function formatTemplateValue(value: TemplateValue): string {
  if (value === null || value === undefined) {
    return MISSING_VALUE_TEXT;
  }

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      return MISSING_VALUE_TEXT;
    }
    return value.toString();
  }

  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      return MISSING_VALUE_TEXT;
    }
    return value.toISOString();
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed.length ? trimmed : MISSING_VALUE_TEXT;
  }

  return MISSING_VALUE_TEXT;
}
