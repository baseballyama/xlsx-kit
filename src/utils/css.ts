// CSS-record → inline-style serializer.
//
// Companion to the `*ToCss` style helpers (fontToCss / fillToCss /
// borderToCss / alignmentToCss / cellStyleToCss): once a caller has
// merged the partials into a `Record<string, string>`, this turns it
// into a stable `style="…"` attribute value for HTML preview.

/**
 * Serialize a CSS-property record to an inline-style declaration string
 * (`prop1: val1; prop2: val2`). Properties are alphabetised so the
 * output is deterministic across runs.
 *
 * - Empty record returns `''`.
 * - Empty-string values are skipped (treat as "unset").
 * - Values containing `;` are dropped — they would terminate the
 *   declaration early and risk attribute-injection in `style="…"`
 *   contexts. Callers should pre-escape user data; this is a
 *   defensive last line.
 *
 * The returned string is suitable for direct interpolation into an
 * HTML `style="…"` attribute *after* the usual attribute-value HTML
 * escaping (no `&` / `"` injection here — this only guards against
 * stray semicolons).
 */
export function cssRecordToInlineStyle(record: Record<string, string> | undefined): string {
  if (!record) return '';
  const parts: string[] = [];
  for (const key of Object.keys(record).sort()) {
    const value = record[key];
    if (value === undefined || value === '') continue;
    if (value.includes(';')) continue;
    parts.push(`${key}: ${value}`);
  }
  return parts.join('; ');
}
