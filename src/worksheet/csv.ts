// CSV serialiser for a worksheet range.
//
// Companion to readRangeAsObjects / writeRangeFromObjects: when the
// caller wants the *raw* delimited form (e.g. for download / paste into
// another tool), this is the export side.

import type { CellValue } from '../cell/cell';
import { getRangeValues, type Worksheet } from './worksheet';

const isObjectKind = (v: unknown, kind: string): boolean =>
  v !== null && typeof v === 'object' && (v as { kind?: string }).kind === kind;

/**
 * Coerce a single CellValue to its CSV-field representation. Strings
 * are returned as-is; numbers / booleans become their `String(...)`;
 * Date becomes ISO-8601; rich-text concatenates run text; formulas
 * use their cached value if present, else the formula source. `null`
 * / unsupported variants become `""`.
 */
function cellToCsvField(v: CellValue | null): string {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toISOString();
  if (isObjectKind(v, 'duration')) return String((v as { ms: number }).ms);
  if (isObjectKind(v, 'error')) return (v as { code: string }).code;
  if (isObjectKind(v, 'rich-text')) {
    const runs = (v as { runs: ReadonlyArray<{ readonly text?: string }> }).runs;
    return runs.map((r) => r.text ?? '').join('');
  }
  if (isObjectKind(v, 'formula')) {
    const fv = v as { cachedValue?: number | string | boolean; formula: string };
    return fv.cachedValue !== undefined ? String(fv.cachedValue) : fv.formula;
  }
  return '';
}

/**
 * Quote a CSV field per RFC 4180: wrap in `"` when the field contains
 * the delimiter, the line terminator, or `"`; double any embedded `"`.
 */
function quoteCsvField(field: string, delimiter: string, lineTerminator: string): string {
  if (
    field.includes(delimiter) ||
    field.includes('"') ||
    field.includes('\n') ||
    field.includes('\r') ||
    field.includes(lineTerminator)
  ) {
    return `"${field.replace(/"/g, '""')}"`;
  }
  return field;
}

/**
 * Render a worksheet range as a CSV string. Empty cells become empty
 * fields. Values are coerced via {@link cellToCsvField} (strings as-is,
 * Dates as ISO-8601, rich-text concatenated, formulas use cached value
 * else formula text).
 *
 * Options:
 *   - `delimiter` (default `,`)
 *   - `lineTerminator` (default `\n`)
 *   - `trailingNewline` (default `false`) — append the terminator after
 *     the last row.
 */
export function getRangeAsCsv(
  ws: Worksheet,
  range: string,
  opts: {
    delimiter?: string;
    lineTerminator?: string;
    trailingNewline?: boolean;
  } = {},
): string {
  const delimiter = opts.delimiter ?? ',';
  const lineTerminator = opts.lineTerminator ?? '\n';
  const grid = getRangeValues(ws, range);
  const lines: string[] = [];
  for (const row of grid) {
    const fields = row.map((v) => quoteCsvField(cellToCsvField(v), delimiter, lineTerminator));
    lines.push(fields.join(delimiter));
  }
  let out = lines.join(lineTerminator);
  if (opts.trailingNewline && out.length > 0) out += lineTerminator;
  return out;
}
