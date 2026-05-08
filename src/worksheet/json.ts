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
import { readRangeAsObjects, type Worksheet } from './worksheet';

type JsonValue = string | number | boolean | null;

const cellValueAsJson = (v: CellValue | null): JsonValue => {
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
  const rows = readRangeAsObjects(
    ws,
    range,
    opts.skipEmptyRows ? { skipEmptyRows: true } : {},
  );
  const mapped: Record<string, JsonValue>[] = rows.map((row) => {
    const obj: Record<string, JsonValue> = {};
    for (const [k, v] of Object.entries(row)) obj[k] = cellValueAsJson(v);
    return obj;
  });
  return JSON.stringify(mapped, null, opts.pretty ? 2 : undefined);
}
