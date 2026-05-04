// Worksheet data model. Per docs/plan/04-core-model.md §4.3.
//
// Cells live in a sparse two-level Map (row → col → Cell). The choice
// is deliberate: a workbook with 1 M cells in 1 column shouldn't
// allocate 1 M empty rows, and JSON.stringify with `Map` round-trips
// cleanly via the workbook's `jsonReplacer`. Worksheets are mutable
// for hot-path performance — see docs/plan/01-architecture.md §5.1.

import type { CellValue } from '../cell/cell';
import { type Cell, makeCell } from '../cell/cell';
import { columnIndexFromLetter, MAX_COL, MAX_ROW } from '../utils/coordinate';
import { OpenXmlSchemaError } from '../utils/exceptions';

export interface Worksheet {
  title: string;
  /** Sparse store: row index → (col index → Cell). */
  rows: Map<number, Map<number, Cell>>;
  /**
   * Highest row index touched by `appendRow`; used to keep appendRow
   * O(1) without re-scanning the row map. Direct setCell / deleteCell
   * may move the actual maximum elsewhere.
   */
  _appendRowCursor: number;
}

/** Build a Worksheet shell. */
export function makeWorksheet(title: string): Worksheet {
  if (typeof title !== 'string' || title.length === 0) {
    throw new OpenXmlSchemaError('Worksheet title must be a non-empty string');
  }
  return {
    title,
    rows: new Map(),
    _appendRowCursor: 0,
  };
}

const validateRowCol = (row: number, col: number): void => {
  if (!Number.isInteger(row) || row < 1 || row > MAX_ROW) {
    throw new OpenXmlSchemaError(`Worksheet row ${row} out of range [1, ${MAX_ROW}]`);
  }
  if (!Number.isInteger(col) || col < 1 || col > MAX_COL) {
    throw new OpenXmlSchemaError(`Worksheet col ${col} out of range [1, ${MAX_COL}]`);
  }
};

/** Resolve a 1-based or "A1" coordinate; returns the populated Cell or undefined. */
export function getCell(ws: Worksheet, row: number, col: number): Cell | undefined {
  return ws.rows.get(row)?.get(col);
}

/**
 * Create or update a Cell at (row, col). Existing cells keep their
 * styleId / hyperlinkId / commentId unless explicitly overridden.
 */
export function setCell(ws: Worksheet, row: number, col: number, value: CellValue = null, styleId?: number): Cell {
  validateRowCol(row, col);
  let rowMap = ws.rows.get(row);
  if (rowMap === undefined) {
    rowMap = new Map<number, Cell>();
    ws.rows.set(row, rowMap);
  }
  let cell = rowMap.get(col);
  if (cell === undefined) {
    cell = makeCell(row, col, value, styleId ?? 0);
    rowMap.set(col, cell);
  } else {
    cell.value = value;
    if (styleId !== undefined) cell.styleId = styleId;
  }
  if (row > ws._appendRowCursor) ws._appendRowCursor = row;
  return cell;
}

/** Delete a single cell from the sheet. Empty rows are pruned. */
export function deleteCell(ws: Worksheet, row: number, col: number): void {
  const rowMap = ws.rows.get(row);
  if (rowMap === undefined) return;
  rowMap.delete(col);
  if (rowMap.size === 0) ws.rows.delete(row);
}

/**
 * Append a row of values starting at the next empty row. Returns the
 * row index (1-based). Mirrors openpyxl's `Worksheet.append`. `null`
 * / `undefined` entries leave the cell empty.
 */
export function appendRow(ws: Worksheet, values: ReadonlyArray<CellValue | undefined>): number {
  const row = ws._appendRowCursor + 1;
  for (let i = 0; i < values.length; i++) {
    const value = values[i];
    if (value === undefined || value === null) continue;
    setCell(ws, row, i + 1, value);
  }
  // Even if every value is empty, advance the cursor so the next call
  // doesn't overwrite this row's would-be position.
  ws._appendRowCursor = row;
  return row;
}

export interface IterRowsOptions {
  minRow?: number;
  maxRow?: number;
  minCol?: number;
  maxCol?: number;
  /** Yield cell values instead of Cell objects. */
  valuesOnly?: boolean;
}

/**
 * Iterate the populated cells row-by-row in ascending order. Empty
 * rows are skipped (no `[]` yielded). Within a row, cells are sorted
 * by column index ascending.
 */
export function* iterRows(ws: Worksheet, opts: IterRowsOptions = {}): IterableIterator<Cell[]> {
  const { minRow = 1, maxRow = MAX_ROW, minCol = 1, maxCol = MAX_COL } = opts;
  const rowKeys = [...ws.rows.keys()].filter((r) => r >= minRow && r <= maxRow).sort((a, b) => a - b);
  for (const r of rowKeys) {
    const rowMap = ws.rows.get(r);
    if (rowMap === undefined) continue;
    const cols = [...rowMap.keys()].filter((c) => c >= minCol && c <= maxCol).sort((a, b) => a - b);
    if (cols.length === 0) continue;
    const out: Cell[] = [];
    for (const c of cols) {
      const cell = rowMap.get(c);
      if (cell !== undefined) out.push(cell);
    }
    yield out;
  }
}

/** Same as `iterRows` but yields each cell's `.value`. */
export function* iterValues(ws: Worksheet, opts: IterRowsOptions = {}): IterableIterator<CellValue[]> {
  for (const row of iterRows(ws, opts)) yield row.map((c) => c.value);
}

/** Effective max row index based on populated cells (0 when empty). */
export function getMaxRow(ws: Worksheet): number {
  let m = 0;
  for (const r of ws.rows.keys()) if (r > m) m = r;
  return m;
}

/** Effective max column index based on populated cells (0 when empty). */
export function getMaxCol(ws: Worksheet): number {
  let m = 0;
  for (const rowMap of ws.rows.values()) {
    for (const c of rowMap.keys()) if (c > m) m = c;
  }
  return m;
}

/** Total populated cell count. */
export function countCells(ws: Worksheet): number {
  let n = 0;
  for (const rowMap of ws.rows.values()) n += rowMap.size;
  return n;
}

/** Resolve an "A1" coordinate to a numeric (col, row) pair on the sheet. */
export function setCellByCoord(ws: Worksheet, coord: string, value?: CellValue, styleId?: number): Cell {
  const m = /^([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(coord);
  if (m === null) {
    throw new OpenXmlSchemaError(`setCellByCoord: invalid coordinate "${coord}"`);
  }
  // biome-ignore lint/style/noNonNullAssertion: matched regex guarantees groups
  const col = columnIndexFromLetter(m[1]!);
  // biome-ignore lint/style/noNonNullAssertion: matched regex guarantees groups
  const row = Number.parseInt(m[2]!, 10);
  return setCell(ws, row, col, value, styleId);
}

/** Convenience getter accepting an "A1" coordinate. */
export function getCellByCoord(ws: Worksheet, coord: string): Cell | undefined {
  const m = /^([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(coord);
  if (m === null) return undefined;
  // biome-ignore lint/style/noNonNullAssertion: matched regex
  const col = columnIndexFromLetter(m[1]!);
  // biome-ignore lint/style/noNonNullAssertion: matched regex
  const row = Number.parseInt(m[2]!, 10);
  return getCell(ws, row, col);
}
