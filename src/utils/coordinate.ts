// Worksheet coordinate utilities. Mirrors openpyxl/openpyxl/utils/cell.py.
//
// Per docs/plan/03-foundations.md §7.1 and docs/plan/01-architecture.md
// §7.4 these functions are on the worksheet read/write hot path
// (millions of calls when streaming a sheet) so the implementations
// stay branch-light, regex-based only at the entry points, with bounded
// Map caches for the bidirectional column letter <-> index mapping.

import { OpenXmlSchemaError } from './exceptions';

/** Maximum column index Excel accepts (XFD). */
export const MAX_COL = 16384;
/** Maximum row index Excel accepts. */
export const MAX_ROW = 1048576;

// ---- column letter <-> 1-based index ---------------------------------------

const indexByLetter = new Map<string, number>();
const letterByIndex = new Map<number, string>();

/**
 * 1-based column index → spreadsheet column letter ("A", "Z", "AA",
 * "XFD"). Throws OpenXmlSchemaError when out of range.
 */
export function columnLetterFromIndex(n: number): string {
  const cached = letterByIndex.get(n);
  if (cached !== undefined) return cached;
  if (!Number.isInteger(n) || n < 1 || n > MAX_COL) {
    throw new OpenXmlSchemaError(`column index ${n} is out of range [1, ${MAX_COL}]`);
  }
  let m = n;
  let out = '';
  while (m > 0) {
    m -= 1; // shift to 0-based for the modulo
    out = String.fromCharCode(65 + (m % 26)) + out;
    m = Math.floor(m / 26);
  }
  letterByIndex.set(n, out);
  return out;
}

/**
 * Column letter → 1-based column index. Case-insensitive but at most
 * 3 letters (the spec ceiling). Throws on empty / non-A-Z / over-range.
 */
export function columnIndexFromLetter(letter: string): number {
  const cached = indexByLetter.get(letter);
  if (cached !== undefined) return cached;
  if (letter.length === 0 || letter.length > 3) {
    throw new OpenXmlSchemaError(`column letter "${letter}" is empty or too long`);
  }
  let n = 0;
  for (let i = 0; i < letter.length; i++) {
    const c = letter.charCodeAt(i);
    let v: number;
    if (c >= 65 && c <= 90)
      v = c - 64; // 'A' = 65
    else if (c >= 97 && c <= 122)
      v = c - 96; // 'a' = 97
    else throw new OpenXmlSchemaError(`column letter "${letter}" contains non-letter char`);
    n = n * 26 + v;
  }
  if (n < 1 || n > MAX_COL) {
    throw new OpenXmlSchemaError(`column letter "${letter}" expands to out-of-range index ${n}`);
  }
  // Normalise the cache key to upper-case so 'a' / 'A' share a slot.
  const key = letter.toUpperCase();
  indexByLetter.set(key, n);
  if (key !== letter) indexByLetter.set(letter, n);
  return n;
}

// ---- coordinate parsing ----------------------------------------------------

/** A single-cell coordinate split into its letter and 1-based row. */
export interface CellCoordinate {
  column: string;
  row: number;
}

/** A single-cell coordinate split into 1-based numeric (col, row). */
export interface CellCoordinateNumeric {
  col: number;
  row: number;
}

/** A 1-based rectangular boundary. */
export interface CellRangeBoundaries {
  minCol: number;
  minRow: number;
  maxCol: number;
  maxRow: number;
}

const COORD_RE = /^[$]?([A-Za-z]{1,3})[$]?([1-9][0-9]*)$/;
const COL_RANGE_RE = /^[$]?([A-Za-z]{1,3}):[$]?([A-Za-z]{1,3})$/;
const ROW_RANGE_RE = /^[$]?([1-9][0-9]*):[$]?([1-9][0-9]*)$/;
const SHEET_RANGE_RE = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/;

/**
 * Parse a single-cell coordinate string ("A1", "$XFD$1048576") into
 * its column letter (always uppercased) and 1-based row.
 */
export function coordinateFromString(coord: string): CellCoordinate {
  const m = COORD_RE.exec(coord);
  if (m === null) throw new OpenXmlSchemaError(`coordinateFromString: invalid coordinate "${coord}"`);
  // biome-ignore lint/style/noNonNullAssertion: regex with two required groups
  const column = m[1]!.toUpperCase();
  // biome-ignore lint/style/noNonNullAssertion: regex with two required groups
  const row = Number.parseInt(m[2]!, 10);
  if (row < 1 || row > MAX_ROW) {
    throw new OpenXmlSchemaError(`coordinateFromString: row ${row} out of range`);
  }
  // Validate column upper bound through the cached helper.
  columnIndexFromLetter(column);
  return { column, row };
}

/**
 * Same as {@link coordinateFromString} but returning the column as its
 * 1-based numeric index. Thin convenience for the worksheet read path.
 */
export function coordinateToTuple(coord: string): CellCoordinateNumeric {
  const c = coordinateFromString(coord);
  return { col: columnIndexFromLetter(c.column), row: c.row };
}

/** Compose `"A1"` from a 1-based (col, row). */
export function tupleToCoordinate(col: number, row: number): string {
  if (!Number.isInteger(row) || row < 1 || row > MAX_ROW) {
    throw new OpenXmlSchemaError(`tupleToCoordinate: row ${row} out of range`);
  }
  return `${columnLetterFromIndex(col)}${row}`;
}

/**
 * Parse "A1:B5" / "A:A" / "1:1" / single-cell into 1-based
 * (minCol, minRow, maxCol, maxRow). Whole-column ranges fill rows to
 * [1, MAX_ROW]; whole-row ranges fill cols to [1, MAX_COL].
 */
export function rangeBoundaries(range: string): CellRangeBoundaries {
  const trimmed = range.trim();
  if (trimmed.length === 0) throw new OpenXmlSchemaError('rangeBoundaries: empty range');

  const colon = trimmed.indexOf(':');
  if (colon < 0) {
    // Single-cell shorthand: "A1" → A1:A1.
    const c = coordinateFromString(trimmed);
    const col = columnIndexFromLetter(c.column);
    return { minCol: col, minRow: c.row, maxCol: col, maxRow: c.row };
  }

  const left = trimmed.slice(0, colon);
  const right = trimmed.slice(colon + 1);

  const colsOnly = COL_RANGE_RE.exec(trimmed);
  if (colsOnly !== null) {
    // biome-ignore lint/style/noNonNullAssertion: matched regex
    const minCol = columnIndexFromLetter(colsOnly[1]!);
    // biome-ignore lint/style/noNonNullAssertion: matched regex
    const maxCol = columnIndexFromLetter(colsOnly[2]!);
    return {
      minCol: Math.min(minCol, maxCol),
      minRow: 1,
      maxCol: Math.max(minCol, maxCol),
      maxRow: MAX_ROW,
    };
  }

  const rowsOnly = ROW_RANGE_RE.exec(trimmed);
  if (rowsOnly !== null) {
    // biome-ignore lint/style/noNonNullAssertion: matched regex
    const minRow = Number.parseInt(rowsOnly[1]!, 10);
    // biome-ignore lint/style/noNonNullAssertion: matched regex
    const maxRow = Number.parseInt(rowsOnly[2]!, 10);
    if (minRow < 1 || maxRow < 1 || minRow > MAX_ROW || maxRow > MAX_ROW) {
      throw new OpenXmlSchemaError(`rangeBoundaries: row out of range in "${trimmed}"`);
    }
    return {
      minCol: 1,
      minRow: Math.min(minRow, maxRow),
      maxCol: MAX_COL,
      maxRow: Math.max(minRow, maxRow),
    };
  }

  const a = coordinateFromString(left);
  const b = coordinateFromString(right);
  const ac = columnIndexFromLetter(a.column);
  const bc = columnIndexFromLetter(b.column);
  return {
    minCol: Math.min(ac, bc),
    minRow: Math.min(a.row, b.row),
    maxCol: Math.max(ac, bc),
    maxRow: Math.max(a.row, b.row),
  };
}

/** Inverse of {@link rangeBoundaries} for the rectangular case. */
export function boundariesToRangeString(b: CellRangeBoundaries): string {
  const tl = tupleToCoordinate(b.minCol, b.minRow);
  if (b.minCol === b.maxCol && b.minRow === b.maxRow) return tl;
  const br = tupleToCoordinate(b.maxCol, b.maxRow);
  return `${tl}:${br}`;
}

/**
 * Parse a sheet-qualified range ("Sheet1!A1:B5" / "'Quarter 1'!A1").
 * Sheet names with single quotes inside use SQL-style doubling
 * ("'Bob''s Sheet'!A1") — we unescape on the way out.
 */
export function parseSheetRange(input: string): {
  sheet: string;
  range: string;
  bounds: CellRangeBoundaries;
} {
  const m = SHEET_RANGE_RE.exec(input);
  if (m === null) throw new OpenXmlSchemaError(`parseSheetRange: missing "!" delimiter in "${input}"`);
  const quoted = m[1];
  const bare = m[2];
  const range = m[3];
  if (range === undefined) throw new OpenXmlSchemaError(`parseSheetRange: empty range part in "${input}"`);
  const sheet = quoted !== undefined ? quoted.replace(/''/g, "'") : (bare ?? '');
  return { sheet, range, bounds: rangeBoundaries(range) };
}
