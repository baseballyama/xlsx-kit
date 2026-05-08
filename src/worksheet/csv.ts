// CSV serialiser for a worksheet range.
//
// Companion to readRangeAsObjects / writeRangeFromObjects: when the
// caller wants the *raw* delimited form (e.g. for download / paste into
// another tool), this is the export side.

import type { CellValue } from '../cell/cell';
import { boundariesToRangeString } from '../utils/coordinate';
import { getDataExtent, getRangeValues, type Worksheet, writeRange } from './worksheet';

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
 * fields. Values are coerced field-by-field (strings as-is, Dates as
 * ISO-8601, rich-text concatenated, formulas use the cached value when
 * present otherwise the formula text).
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

/**
 * Whole-worksheet shortcut over {@link getRangeAsCsv}: serialises the
 * sheet's data extent (`getDataExtent`) as CSV. Returns `''` for an
 * empty worksheet. All `getRangeAsCsv` opts are forwarded.
 */
export function getWorksheetAsCsv(
  ws: Worksheet,
  opts: {
    delimiter?: string;
    lineTerminator?: string;
    trailingNewline?: boolean;
  } = {},
): string {
  const ext = getDataExtent(ws);
  if (!ext) return '';
  return getRangeAsCsv(ws, boundariesToRangeString(ext), opts);
}

/**
 * RFC 4180 CSV parser. Handles quoted fields, embedded delimiters /
 * newlines / `""`. Accepts both `\n` and `\r\n` line terminators on
 * input regardless of `opts.delimiter`.
 *
 * Returns a 2D `string[][]`. Empty input returns `[]`.
 */
export function parseCsv(input: string, opts: { delimiter?: string } = {}): string[][] {
  const delimiter = opts.delimiter ?? ',';
  if (input.length === 0) return [];
  const out: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  while (i < input.length) {
    const c = input[i];
    if (inQuotes) {
      if (c === '"') {
        if (input[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += c;
      i++;
      continue;
    }
    if (c === '"' && field.length === 0) {
      inQuotes = true;
      i++;
      continue;
    }
    if (c === delimiter) {
      row.push(field);
      field = '';
      i++;
      continue;
    }
    if (c === '\r' && input[i + 1] === '\n') {
      row.push(field);
      out.push(row);
      row = [];
      field = '';
      i += 2;
      continue;
    }
    if (c === '\n' || c === '\r') {
      row.push(field);
      out.push(row);
      row = [];
      field = '';
      i++;
      continue;
    }
    field += c;
    i++;
  }
  // Flush the in-flight row. Don't emit a trailing empty row when the
  // input ended with a clean line terminator (`field === '' && row.length === 0`).
  if (field !== '' || row.length > 0) {
    row.push(field);
    out.push(row);
  }
  return out;
}

const tryCoerce = (s: string): CellValue => {
  if (s === '') return '';
  if (s === 'true') return true;
  if (s === 'false') return false;
  // Excel-friendly numeric coerce: strict integer or decimal, no scientific edge-cases.
  if (/^-?\d+(\.\d+)?$/.test(s)) {
    const n = Number(s);
    if (Number.isFinite(n)) return n;
  }
  return s;
};

/**
 * Inverse of {@link getRangeAsCsv}: parse a CSV string and write the
 * resulting 2D grid to the worksheet starting at `startRef` via
 * {@link writeRange}. Returns the bounding-box (or `undefined` for
 * empty input).
 *
 * Options:
 *   - `delimiter` (default `,`)
 *   - `coerceTypes` (default `false`) — when `true`, parse `"true"` /
 *     `"false"` to booleans and integer / decimal strings to numbers.
 *     Otherwise everything stays as a string.
 */
export function parseCsvToRange(
  ws: Worksheet,
  startRef: string,
  csv: string,
  opts: { delimiter?: string; coerceTypes?: boolean } = {},
): { minRow: number; maxRow: number; minCol: number; maxCol: number } | undefined {
  const grid = parseCsv(csv, { delimiter: opts.delimiter ?? ',' });
  if (grid.length === 0) return undefined;
  const transformed: Array<Array<CellValue | undefined>> = grid.map((row) =>
    row.map((field) => {
      const v = opts.coerceTypes ? tryCoerce(field) : field;
      // Empty fields stay as '' so they don't unintentionally clear cells —
      // writeRange only skips on `undefined` / `null`. Callers wanting the
      // skip semantic should pre-process.
      return v as CellValue | undefined;
    }),
  );
  return writeRange(ws, startRef, transformed);
}
