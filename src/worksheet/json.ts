// Worksheet → JSON renderer. Sibling of worksheetToCsv /
// worksheetToHtml / worksheetToMarkdownTable / worksheetToTextTable
// for the "give me the data as JSON" output side.
//
// The first row of the range becomes object keys (header). Subsequent
// rows become array entries. Cell values are mapped to JSON-safe
// equivalents:
//
//   - Date            → ISO 8601 string (`Date.prototype.toISOString`)
//   - formula         → `cachedValue` if set, else the formula text
//   - duration        → millisecond number (`{ ms }.ms`)
//   - error           → Excel error token (`'#REF!'` etc.)
//   - rich-text       → concatenated plain text of all runs
//   - other primitives pass through verbatim

import {
  type CellValue,
  isDurationValue,
  isErrorValue,
  isFormulaValue,
  isRichTextValue,
} from '../cell/cell';
import { boundariesToRangeString } from '../utils/coordinate';
import { getDataExtent, readRangeAsObjects, type Worksheet, writeRange } from './worksheet';

export type JsonValue = string | number | boolean | null;
export type JsonRow = Record<string, JsonValue>;

export const cellValueAsJson = (v: CellValue | null): JsonValue => {
  if (v === null) return null;
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v;
  if (v instanceof Date) return v.toISOString();
  if (isFormulaValue(v)) {
    if (v.cachedValue !== undefined) return v.cachedValue;
    return v.formula;
  }
  if (isRichTextValue(v)) {
    let s = '';
    for (const r of v.runs) s += r.text;
    return s;
  }
  if (isDurationValue(v)) return v.ms;
  if (isErrorValue(v)) return v.code;
  return null;
};

export interface WorksheetToJsonOptions {
  /** Pretty-print with 2-space indentation. Default: false (single-line). */
  pretty?: boolean;
  /** Drop rows where every column is empty. Forwarded to {@link readRangeAsObjects}. */
  skipEmptyRows?: boolean;
}

/**
 * Serialise a worksheet range as a JSON string. The first row of the
 * range supplies the object keys; subsequent rows become array
 * entries. Returns `'[]'` for a header-only range (no data rows).
 *
 * Cell values are coerced to JSON-safe shapes:
 * `Date → ISO string`, `formula → cachedValue ?? formula text`,
 * `duration → ms number`, `error → token`,
 * `rich-text → concatenated text`. Other primitives pass through.
 *
 * `opts.pretty` applies 2-space indentation. `opts.skipEmptyRows`
 * drops fully-blank rows (forwarded to `readRangeAsObjects`).
 */
export function worksheetToJson(
  ws: Worksheet,
  range: string,
  opts: WorksheetToJsonOptions = {},
): string {
  return JSON.stringify(worksheetRowsAsJson(ws, range, opts), null, opts.pretty ? 2 : undefined);
}

/**
 * Build an array of JSON-safe row objects for a range. Same coercion
 * as {@link worksheetToJson} but returns structured rows instead of
 * a serialised string, so workbook-wide aggregators can `JSON.stringify`
 * once over the combined shape (e.g. {@link getWorkbookAsJsonString}).
 */
export function worksheetRowsAsJson(
  ws: Worksheet,
  range: string,
  opts: WorksheetToJsonOptions = {},
): JsonRow[] {
  const rows = readRangeAsObjects(
    ws,
    range,
    opts.skipEmptyRows ? { skipEmptyRows: true } : {},
  );
  return rows.map((row) => {
    const obj: JsonRow = {};
    for (const [k, v] of Object.entries(row)) obj[k] = cellValueAsJson(v);
    return obj;
  });
}

/**
 * Whole-worksheet shortcut over {@link worksheetToJson}: serialises
 * the sheet's data extent (`getDataExtent`) as JSON. Returns `'[]'`
 * for an empty worksheet (mirrors the CSV / HTML / Markdown / Text
 * shortcut conventions, with `'[]'` chosen so the output stays a
 * valid JSON document).
 */
export function getWorksheetAsJson(ws: Worksheet, opts: WorksheetToJsonOptions = {}): string {
  const ext = getDataExtent(ws);
  if (!ext) return '[]';
  return worksheetToJson(ws, boundariesToRangeString(ext), opts);
}

/**
 * Whole-worksheet variant of {@link worksheetRowsAsJson} (data extent →
 * `JsonRow[]`). Returns `[]` for an empty worksheet. Used by
 * {@link getWorkbookAsJsonString} to assemble a single combined JSON
 * document without re-parsing per-sheet output.
 */
export function getWorksheetRowsAsJson(
  ws: Worksheet,
  opts: WorksheetToJsonOptions = {},
): JsonRow[] {
  const ext = getDataExtent(ws);
  if (!ext) return [];
  return worksheetRowsAsJson(ws, boundariesToRangeString(ext), opts);
}

const ISO_8601_RE =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?(?:Z|[+-]\d{2}:?\d{2})$/;

/**
 * Coerce a single JSON-decoded value back into a {@link CellValue}.
 * Inverse partner of {@link cellValueAsJson}: numbers / booleans / null
 * pass through, strings matching ISO 8601 (`YYYY-MM-DDTHH:MM:SS[.sss][Z|±hh:mm]`)
 * are restored to `Date`, other strings stay strings, and any other
 * shape (object / array) falls back to `String(v)`.
 */
export const cellValueFromJson = (v: unknown): CellValue => {
  if (v === null) return null;
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  if (typeof v === 'string') {
    if (ISO_8601_RE.test(v)) {
      const d = new Date(v);
      if (!Number.isNaN(d.getTime())) return d;
    }
    return v;
  }
  return String(v);
};

export interface ParseJsonToRangeOptions {
  /**
   * Header order to use when writing. Each key becomes a column in the
   * order given, regardless of where it appears in each row object.
   * Defaults to `Object.keys(rows[0])` — the insertion order of the
   * first row.
   */
  keys?: string[];
}

/**
 * Inverse of {@link worksheetToJson}: parse a JSON array of row
 * objects (`[{name: "Alice", age: 30}, …]` — the shape produced by
 * `worksheetToJson`) and write it to the worksheet as a header row
 * plus one data row per array entry, anchored at `topLeft`.
 *
 * `json` may be a JSON string (parsed via `JSON.parse`) or an
 * already-decoded array. Returns the bounding box of the written
 * range (like {@link writeRange}), or `undefined` for an empty
 * array (no header is written when there are no rows).
 *
 * Cell values are coerced via {@link cellValueFromJson}: ISO 8601
 * date strings become `Date`, primitives pass through, and other
 * shapes fall back to `String(v)`. Missing keys in a given row are
 * written as `null`.
 *
 * `opts.keys` overrides the header order; otherwise the first row's
 * own key order is used.
 */
export function parseJsonToRange(
  ws: Worksheet,
  topLeft: string,
  json: string | readonly unknown[],
  opts: ParseJsonToRangeOptions = {},
): { minRow: number; maxRow: number; minCol: number; maxCol: number } | undefined {
  const decoded = typeof json === 'string' ? (JSON.parse(json) as unknown) : json;
  if (!Array.isArray(decoded) || decoded.length === 0) return undefined;
  const rows = decoded as readonly unknown[];
  const firstRow = rows[0];
  if (firstRow === null || typeof firstRow !== 'object' || Array.isArray(firstRow)) {
    throw new TypeError('parseJsonToRange: rows must be JSON objects');
  }
  const keys = opts.keys ?? Object.keys(firstRow as Record<string, unknown>);
  const grid: Array<Array<CellValue | undefined>> = [keys];
  for (const row of rows) {
    if (row === null || typeof row !== 'object' || Array.isArray(row)) {
      throw new TypeError('parseJsonToRange: rows must be JSON objects');
    }
    const obj = row as Record<string, unknown>;
    grid.push(keys.map((k) => (k in obj ? cellValueFromJson(obj[k]) : null)));
  }
  return writeRange(ws, topLeft, grid);
}
