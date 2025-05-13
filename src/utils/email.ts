const EMAIL_RE = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;

export function extractEmail(raw: unknown): string | null {
  if (typeof raw !== 'string') return null;
  const m = raw.match(EMAIL_RE);
  return m?.[ 0 ]?.trim() ?? null;
}

export const renderTemplate = (s: string, vars: Record<string, string>) =>
  s.replace(/\{\{(\w+)\}\}/g, (_, k) => vars[ k ] ?? '');
