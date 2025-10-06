export function slugifyString(value: string): string {
  const trimmed = value.trim();
  const withoutIllegal = trimmed.replace(/[<>:"/\\|?*]/g, '');
  const collapsedWhitespace = withoutIllegal.replace(/\s+/g, ' ').trim();
  const dashed = collapsedWhitespace.replace(/\s+/g, '-');
  const cleaned = dashed.replace(/-+/g, '-');
  return cleaned.toLowerCase() || 'report';
}
