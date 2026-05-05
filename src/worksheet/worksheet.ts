// Worksheet data model. Per docs/plan/04-core-model.md §4.3.
//
// Cells live in a sparse two-level Map (row → col → Cell). The choice
// is deliberate: a workbook with 1 M cells in 1 column shouldn't
// allocate 1 M empty rows, and JSON.stringify with `Map` round-trips
// cleanly via the workbook's `jsonReplacer`. Worksheets are mutable
// for hot-path performance — see docs/plan/01-architecture.md §5.1.

import type { CellValue } from '../cell/cell';
import { type Cell, cellValueAsString, makeCell } from '../cell/cell';
import type { Drawing } from '../drawing/drawing';
import { type Color, makeColor } from '../styles/colors';
import { columnIndexFromLetter, MAX_COL, MAX_ROW } from '../utils/coordinate';
import { OpenXmlSchemaError } from '../utils/exceptions';
import type { AutoFilter } from './auto-filter';
import { type CellRange, parseRange, rangeContainsCell, rangesOverlap, rangeToString } from './cell-range';
import type { LegacyComment } from './comments';
import { makeLegacyComment } from './comments';
import type { ConditionalFormatting } from './conditional-formatting';
import type { DataValidation } from './data-validations';
import { type ColumnDimension, makeColumnDimension, makeRowDimension, type RowDimension } from './dimensions';
import type { DataConsolidate } from './data-consolidate';
import type { ScenarioList } from './scenarios';
import type { CellWatch, IgnoredError } from './errors';
import type { HeaderFooter, PageBreak, PageMargins, PageSetup, PrintOptions } from './page-setup';
import type { WorksheetPhoneticProperties } from './phonetic';
import { makeSheetProperties, type SheetProperties } from './properties';
import type { SheetProtection } from './protection';
import type { ProtectedRange } from './protected-ranges';
import type { SortState } from './sort-state';
import type { WebPublishItem, WorksheetCustomProperty } from './web-publish';
import { type Hyperlink, makeHyperlink } from './hyperlinks';
import type { TableDefinition } from './table';
import { freezePaneRef, makeFreezePane, makeSheetView, type SheetView } from './views';

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
  /**
   * Merged cell ranges. The top-left cell holds the value; the rest are
   * mostly invisible to Excel until unmerge restores them. We persist the
   * list as plain CellRange[] so mergeCells / unmergeCells can mutate it
   * without rebuilding any helper structures.
   */
  mergedCells: CellRange[];
  /**
   * Per-bookView display settings. Most workbooks have exactly one view.
   * The list stays empty until something — read or API — populates it; a
   * lone default view doesn't earn its keep on the wire.
   */
  views: SheetView[];
  /**
   * Per-column metadata keyed by the entry's `min` index. The value's
   * `min`/`max` may cover a multi-column run; iteration over the map
   * yields one entry per run, in `min`-ascending order.
   */
  columnDimensions: Map<number, ColumnDimension>;
  /** Per-row metadata keyed by 1-based row index. */
  rowDimensions: Map<number, RowDimension>;
  /** Default column width (characters) when not overridden by a column dimension. */
  defaultColumnWidth?: number;
  /** Default row height (points) when not overridden by a row dimension. */
  defaultRowHeight?: number;
  /**
   * Highest outline depth used among `rowDimensions`. Excel uses this to
   * size the outline button strip on the left of the row numbers. Auto-
   * computed by the writer when undefined — set explicitly to override.
   */
  outlineLevelRow?: number;
  /** Highest outline depth used among `columnDimensions`. Auto-computed by the writer when undefined. */
  outlineLevelCol?: number;
  /** "Custom row heights present" hint — Excel uses this to skip default-height rendering. */
  customHeight?: boolean;
  /** "Show rows of zero height as one row" — for hidden rows the outline collapse uses this. */
  zeroHeight?: boolean;
  /** Apply a thick top border to every row by default. */
  thickTop?: boolean;
  /** Apply a thick bottom border to every row by default. */
  thickBottom?: boolean;
  /** Excel's "base column width" (characters) — defaults to 8 when unset. */
  baseColWidth?: number;
  /** Hyperlinks. External URLs round-trip via worksheet rels; internal jumps stay inline. */
  hyperlinks: Hyperlink[];
  /** Data validation entries. */
  dataValidations: DataValidation[];
  /** AutoFilter — at most one per sheet. Excel reuses the `_xlnm._FilterDatabase` defined name. */
  autoFilter?: AutoFilter;
  /** Excel Table objects. Each lives in its own xl/tables/tableN.xml part. */
  tables: TableDefinition[];
  /** Legacy comments. Persisted as `xl/commentsN.xml` + a placeholder VML drawing. */
  legacyComments: LegacyComment[];
  /** Conditional formatting blocks. */
  conditionalFormatting: ConditionalFormatting[];
  /**
   * `<sheetPr>` properties — VBA codeName, tab strip color, outline /
   * page-setup defaults, etc. Top-level `<sheetPr>` lives just before
   * `<dimension>` per ECMA-376 ordering.
   */
  sheetProperties?: SheetProperties;
  /**
   * Sheet-protection state. When `sheet=true` Excel locks the sheet
   * against edits (subject to the per-action allow flags here).
   * Password hashing helpers come later — for now `saltValue` /
   * `spinCount` / `algorithmName` / `hashValue` round-trip verbatim.
   */
  sheetProtection?: SheetProtection;
  /**
   * `<protectedRanges>` — per-range edit-allowance overrides used when
   * the sheet is otherwise protected (Review → Allow Edit Ranges).
   */
  protectedRanges: ProtectedRange[];
  /**
   * `<sortState>` — last-applied sort criteria. Excel persists this so
   * the rows come back in the same order after a save/load cycle.
   */
  sortState?: SortState;
  /**
   * `<picture r:id="…"/>` — sheet background image (Page Layout →
   * Background). The rId points at a media part registered in the
   * worksheet rels (preserved via the existing relsExtras machinery).
   */
  backgroundPictureRId?: string;
  /**
   * `<legacyDrawingHF r:id="…"/>` — VML drawing used for header/footer
   * background images on print. Parallel to legacyDrawing (which
   * carries comment markers); the rels link rides relsExtras.
   */
  legacyDrawingHFRId?: string;
  /**
   * `<smartTags>` — per-cell smart-tag annotations (Excel 2003 era).
   * Pairs with the workbook-level smartTagTypes registry.
   */
  smartTags: import('./smart-tags').CellSmartTags[];
  /**
   * `<customSheetViews>` — saved per-user view presets for this
   * worksheet (Excel's "Custom Views" feature). Each entry snapshots
   * zoom / gridline / formula / heading toggles plus its own page-
   * setup block and break list.
   */
  customSheetViews: import('./custom-sheet-views').CustomSheetView[];
  /**
   * `<oleObjects>` — embedded OLE objects (linked Word documents,
   * Equation editor formulas, etc.). The objectPr child is round-
   * tripped verbatim as an XmlNode.
   */
  oleObjects: import('./ole-objects').OleObject[];
  /**
   * `<controls>` — form controls (checkboxes / list boxes / spin
   * buttons placed via the Developer tab). The controlPr child is
   * round-tripped verbatim.
   */
  controls: import('./ole-objects').FormControl[];
  /** `<printOptions>` — gridlines, headings, horizontal/vertical centering on the printed page. */
  printOptions?: PrintOptions;
  /** `<pageMargins>` — six required margins in inches. */
  pageMargins?: PageMargins;
  /** `<pageSetup>` — paper size / orientation / scale / fitToPage / DPI etc. */
  pageSetup?: PageSetup;
  /** `<headerFooter>` — odd/even/first header + footer mini-format strings + flags. */
  headerFooter?: HeaderFooter;
  /** Manual horizontal page breaks (`<rowBreaks>`). Each entry's `id` is the row above which a new page begins. */
  rowBreaks: PageBreak[];
  /** Manual vertical page breaks (`<colBreaks>`). Each entry's `id` is the column to the left of which a new page begins. */
  colBreaks: PageBreak[];
  /**
   * Worksheet-level `<customProperties>` — per-sheet user metadata that
   * SharePoint workflows attach (separate from the workbook-level
   * `docProps/custom.xml` part). The `rId` points at a Custom XML part
   * registered in the worksheet rels (already preserved via `relsExtras`).
   */
  customProperties: WorksheetCustomProperty[];
  /** `<webPublishItems>` — Excel 2007's "Publish to web" entries. Almost always empty in modern files. */
  webPublishItems: WebPublishItem[];
  /**
   * `<phoneticPr>` — East-Asian furigana rendering hints (font index +
   * IME conversion mode + alignment). Common in Japanese workbooks.
   */
  phoneticPr?: WorksheetPhoneticProperties;
  /**
   * `<dataConsolidate>` — config for Data → Consolidate. Carries the
   * aggregation function and the source-range list.
   */
  dataConsolidate?: DataConsolidate;
  /** `<scenarios>` — the Scenario Manager's saved input-cell overrides. */
  scenarios?: ScenarioList;
  /** Cells pinned in Excel's Watch Window (`<cellWatches><cellWatch r="…"/></cellWatches>`). */
  cellWatches: CellWatch[];
  /**
   * Per-region "ignore this error class" rules (`<ignoredErrors>`).
   * Suppresses the small green-triangle warning for the listed checks.
   */
  ignoredErrors: IgnoredError[];
  /**
   * Spreadsheet drawing — at most one per worksheet. Hosts charts /
   * pictures / shapes. Persisted as `xl/drawings/drawingN.xml` plus a
   * worksheet-rels entry; on the wire the worksheet body emits
   * `<drawing r:id>`.
   */
  drawing?: Drawing;
  /**
   * Per-sheet rels entries we don't model (pivotTable / queryTable /
   * slicer / printerSettings / customProperty / oleObject etc.). Re-emitted
   * verbatim so Excel still resolves the captured passthrough parts after
   * a round-trip.
   */
  relsExtras?: ReadonlyArray<{ id: string; type: string; target: string }>;
  /**
   * Top-level `<worksheet>` children we don't model — `<sheetPr>`,
   * `<printOptions>`, `<pageMargins>`, `<pageSetup>`, `<headerFooter>`,
   * `<rowBreaks>`, `<colBreaks>`, `<oleObjects>`, `<controls>`,
   * `<picture>`, `<extLst>`, etc. Captured as XmlNodes; the writer emits
   * them in two anchored slots so common ECMA-376 ordering survives a
   * round-trip even though we don't track every position individually:
   * - `beforeSheetData` → emitted before our `<dimension>` (typical for
   *   `<sheetPr>`).
   * - `afterSheetData` → emitted between our `<hyperlinks>` and
   *   `<drawing>` block, which lands page setup / extLst / oleObjects in
   *   roughly the right place. Excel reads back regardless of strict
   *   ECMA position; openpyxl-emitted files round-trip cleanly.
   */
  bodyExtras?: {
    beforeSheetData: import('../xml/tree').XmlNode[];
    afterSheetData: import('../xml/tree').XmlNode[];
  };
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
    mergedCells: [],
    views: [],
    columnDimensions: new Map(),
    rowDimensions: new Map(),
    hyperlinks: [],
    dataValidations: [],
    tables: [],
    legacyComments: [],
    conditionalFormatting: [],
    cellWatches: [],
    ignoredErrors: [],
    rowBreaks: [],
    colBreaks: [],
    customProperties: [],
    webPublishItems: [],
    protectedRanges: [],
    smartTags: [],
    customSheetViews: [],
    oleObjects: [],
    controls: [],
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

/**
 * Bounding-box of the populated cells: `{ minRow, maxRow, minCol,
 * maxCol }` covering every cell in `ws.rows`. Returns `undefined`
 * when the sheet is empty. Walks the sparse store once.
 */
export function getDataExtent(
  ws: Worksheet,
): { minRow: number; maxRow: number; minCol: number; maxCol: number } | undefined {
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let minCol = Number.POSITIVE_INFINITY;
  let maxCol = 0;
  let touched = false;
  for (const [r, rowMap] of ws.rows) {
    if (rowMap.size === 0) continue;
    if (r < minRow) minRow = r;
    if (r > maxRow) maxRow = r;
    for (const c of rowMap.keys()) {
      if (c < minCol) minCol = c;
      if (c > maxCol) maxCol = c;
    }
    touched = true;
  }
  if (!touched) return undefined;
  return { minRow, maxRow, minCol, maxCol };
}

/**
 * Same as {@link getDataExtent} but returns the canonical `"A1:E10"`
 * range string for the bounding box, or `undefined` when the sheet
 * is empty.
 */
export function getDataExtentRef(ws: Worksheet): string | undefined {
  const ext = getDataExtent(ws);
  if (!ext) return undefined;
  return rangeToString({
    minRow: ext.minRow,
    maxRow: ext.maxRow,
    minCol: ext.minCol,
    maxCol: ext.maxCol,
  });
}

/**
 * Iterate every populated cell, yielding those for which `predicate`
 * returns true. Iteration order is row-then-column ascending. Cells
 * whose `.value === null` (empty placeholders carrying only style or
 * comment metadata) are still visited.
 */
export function* findCells(
  ws: Worksheet,
  predicate: (c: Cell) => boolean,
): IterableIterator<Cell> {
  const rowKeys = [...ws.rows.keys()].sort((a, b) => a - b);
  for (const r of rowKeys) {
    const rowMap = ws.rows.get(r);
    if (!rowMap) continue;
    const cols = [...rowMap.keys()].sort((a, b) => a - b);
    for (const c of cols) {
      const cell = rowMap.get(c);
      if (cell !== undefined && predicate(cell)) yield cell;
    }
  }
}

/** First populated cell satisfying `predicate`, or `undefined`. */
export function findFirstCell(
  ws: Worksheet,
  predicate: (c: Cell) => boolean,
): Cell | undefined {
  for (const cell of findCells(ws, predicate)) return cell;
  return undefined;
}

/**
 * Find-and-replace across populated string cells. `search` matches
 * either an exact-string equal (when given a string) or every cell
 * whose value satisfies the predicate (when given a function).
 * `replacement` is the new value for each match. Returns the count
 * of cells changed. Non-string-valued cells are skipped when
 * `search` is a string; predicate-based searches see every cell.
 */
export function replaceCellValues(
  ws: Worksheet,
  search: string | ((value: CellValue, cell: Cell) => boolean),
  replacement: CellValue,
): number {
  let n = 0;
  const matchFn =
    typeof search === 'string'
      ? (v: CellValue) => typeof v === 'string' && v === search
      : (v: CellValue, c: Cell) => search(v, c);
  for (const rowMap of ws.rows.values()) {
    for (const cell of rowMap.values()) {
      if (matchFn(cell.value, cell)) {
        cell.value = replacement;
        n++;
      }
    }
  }
  return n;
}

/**
 * Iterate the populated cells inside a rectangular range. Cells that
 * don't exist in the sparse store are skipped (no auto-allocate). Use
 * {@link applyToRange} when you need every coordinate visited
 * regardless of population.
 */
export function* getCellsInRange(ws: Worksheet, range: string): IterableIterator<Cell> {
  const { minRow, maxRow, minCol, maxCol } = parseRange(range);
  for (let r = minRow; r <= maxRow; r++) {
    const rowMap = ws.rows.get(r);
    if (!rowMap) continue;
    for (let col = minCol; col <= maxCol; col++) {
      const cell = rowMap.get(col);
      if (cell !== undefined) yield cell;
    }
  }
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

// ---- merged cells ---------------------------------------------------------

const toCellRange = (refOrRange: string | CellRange): CellRange =>
  typeof refOrRange === 'string' ? parseRange(refOrRange) : refOrRange;

/**
 * Merge a range. The top-left cell keeps its value; every other cell in
 * the range is dropped from `ws.rows` so the on-wire `<sheetData>` won't
 * carry phantom cells underneath the merge. Mirrors openpyxl's
 * `MergedCellRange.format()`. Idempotent for an identical range, throws
 * when the range overlaps an existing merge.
 */
export function mergeCells(ws: Worksheet, refOrRange: string | CellRange): CellRange {
  const range = toCellRange(refOrRange);
  for (const existing of ws.mergedCells) {
    if (rangeToString(existing) === rangeToString(range)) return existing;
    if (rangesOverlap(existing, range)) {
      throw new OpenXmlSchemaError(
        `mergeCells: range ${rangeToString(range)} overlaps existing merged range ${rangeToString(existing)}`,
      );
    }
  }
  // Drop every cell except the top-left from the sparse store.
  for (let r = range.minRow; r <= range.maxRow; r++) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      if (r === range.minRow && c === range.minCol) continue;
      ws.rows.get(r)?.delete(c);
      const row = ws.rows.get(r);
      if (row && row.size === 0) ws.rows.delete(r);
    }
  }
  ws.mergedCells.push(range);
  return range;
}

/** Drop a previously-merged range. No-op if the range isn't registered. */
export function unmergeCells(ws: Worksheet, refOrRange: string | CellRange): boolean {
  const target = rangeToString(toCellRange(refOrRange));
  const idx = ws.mergedCells.findIndex((r) => rangeToString(r) === target);
  if (idx < 0) return false;
  ws.mergedCells.splice(idx, 1);
  return true;
}

/** Read-only iterator over the worksheet's merged ranges. */
export function getMergedCells(ws: Worksheet): ReadonlyArray<CellRange> {
  return ws.mergedCells;
}

/** True iff (row, col) sits inside any merged range — top-left included. */
export function isMergedCell(ws: Worksheet, row: number, col: number): boolean {
  for (const range of ws.mergedCells) {
    if (rangeContainsCell(range, row, col)) return true;
  }
  return false;
}

// ---- views / freezePanes --------------------------------------------------

/** Lazily get-or-create the primary SheetView so view-mutating helpers don't have to branch. */
const ensurePrimaryView = (ws: Worksheet): SheetView => {
  let view = ws.views[0];
  if (!view) {
    view = makeSheetView();
    ws.views.push(view);
  }
  return view;
};

/**
 * Freeze rows / columns above + left of `topLeftRef` ("B2" → 1 row + 1 col).
 * Pass `undefined` to clear any existing freeze. Targets the workbook's
 * primary SheetView (`ws.views[0]`); creates one if absent.
 */
export function setFreezePanes(ws: Worksheet, topLeftRef: string | undefined): void {
  if (topLeftRef === undefined) {
    if (ws.views[0]) delete ws.views[0].pane;
    return;
  }
  const view = ensurePrimaryView(ws);
  view.pane = makeFreezePane(topLeftRef);
}

/** Inverse of {@link setFreezePanes}; returns the top-left ref or undefined when no freeze is active. */
export function getFreezePanes(ws: Worksheet): string | undefined {
  const view = ws.views[0];
  if (!view) return undefined;
  return freezePaneRef(view);
}

import { columnLetterFromIndex } from '../utils/coordinate';

/**
 * Freeze the top `count` rows. Equivalent to
 * `setFreezePanes(ws, "A${count + 1}")` — Excel's "Freeze Top Row"
 * is `freezeRows(ws, 1)`.
 */
export function freezeRows(ws: Worksheet, count: number): void {
  if (!Number.isInteger(count) || count < 1) {
    throw new OpenXmlSchemaError(`freezeRows: count must be a positive integer; got ${count}`);
  }
  setFreezePanes(ws, `A${count + 1}`);
}

/**
 * Freeze the leftmost `count` columns. Equivalent to
 * `setFreezePanes(ws, "${columnLetter(count + 1)}1")` — Excel's
 * "Freeze First Column" is `freezeColumns(ws, 1)`.
 */
export function freezeColumns(ws: Worksheet, count: number): void {
  if (!Number.isInteger(count) || count < 1) {
    throw new OpenXmlSchemaError(`freezeColumns: count must be a positive integer; got ${count}`);
  }
  setFreezePanes(ws, `${columnLetterFromIndex(count + 1)}1`);
}

/** Freeze both top `rows` rows AND left `cols` columns. */
export function freezePanes(ws: Worksheet, rows: number, cols: number): void {
  if (!Number.isInteger(rows) || rows < 1) {
    throw new OpenXmlSchemaError(`freezePanes: rows must be a positive integer; got ${rows}`);
  }
  if (!Number.isInteger(cols) || cols < 1) {
    throw new OpenXmlSchemaError(`freezePanes: cols must be a positive integer; got ${cols}`);
  }
  setFreezePanes(ws, `${columnLetterFromIndex(cols + 1)}${rows + 1}`);
}

/** Drop the freeze pane on the primary view. */
export const unfreezePanes = (ws: Worksheet): void => {
  setFreezePanes(ws, undefined);
};

// ---- sheet view display helpers -------------------------------------------

/** Lazily get-or-create `ws.sheetProperties` so tab-color helpers don't have to branch. */
const ensureSheetProperties = (ws: Worksheet): SheetProperties => {
  if (!ws.sheetProperties) ws.sheetProperties = makeSheetProperties();
  return ws.sheetProperties;
};

const colorFrom = (input: string | Partial<Color>): Color =>
  typeof input === 'string' ? makeColor({ rgb: input }) : makeColor(input);

/**
 * Set the sheet tab strip colour. Accepts either a hex string
 * (`"FF0070C0"`) or a partial `Color` object (`{ theme: 4, tint: 0.4 }`).
 */
export function setSheetTabColor(ws: Worksheet, color: string | Partial<Color>): Color {
  const c = colorFrom(color);
  ensureSheetProperties(ws).tabColor = c;
  return c;
}

/** Drop the sheet tab strip colour. */
export function removeSheetTabColor(ws: Worksheet): void {
  if (!ws.sheetProperties) return;
  delete (ws.sheetProperties as { tabColor?: Color }).tabColor;
}

/** Toggle gridline display on the primary SheetView. */
export function setShowGridLines(ws: Worksheet, show: boolean): void {
  ensurePrimaryView(ws).showGridLines = show;
}

/** Toggle row + column header display on the primary SheetView. */
export function setShowRowColHeaders(ws: Worksheet, show: boolean): void {
  ensurePrimaryView(ws).showRowColHeaders = show;
}

/** Toggle "Show Formulas" mode on the primary SheetView. */
export function setShowFormulas(ws: Worksheet, show: boolean): void {
  ensurePrimaryView(ws).showFormulas = show;
}

/** Toggle "Show a zero in cells that have a zero value" on the primary SheetView. */
export function setShowZeros(ws: Worksheet, show: boolean): void {
  ensurePrimaryView(ws).showZeros = show;
}

/** Toggle right-to-left layout on the primary SheetView. */
export function setRightToLeft(ws: Worksheet, rtl: boolean): void {
  ensurePrimaryView(ws).rightToLeft = rtl;
}

/**
 * Set the zoom scale (percent) on the primary SheetView. Excel accepts
 * integer percentages in `[10, 400]`.
 */
export function setSheetZoom(ws: Worksheet, scale: number): void {
  if (!Number.isInteger(scale) || scale < 10 || scale > 400) {
    throw new OpenXmlSchemaError(`setSheetZoom: scale must be an integer in [10, 400]; got ${scale}`);
  }
  ensurePrimaryView(ws).zoomScale = scale;
}

/** Switch the sheet view between Excel's "Normal" / "Page Break Preview" / "Page Layout" modes. */
export function setSheetViewMode(ws: Worksheet, mode: 'normal' | 'pageBreakPreview' | 'pageLayout'): void {
  ensurePrimaryView(ws).view = mode;
}

/**
 * Set the active cell on the primary SheetView. The active cell is
 * the one Excel highlights with the dark border when the sheet is
 * opened. Updates the existing Selection; creates one if missing.
 * Pass an "A1"-style ref.
 */
export function setActiveCell(ws: Worksheet, ref: string): void {
  const view = ensurePrimaryView(ws);
  const selection = view.selection ?? {};
  // Excel typically also sets sqref to the same cell when a single
  // cell is the active one. Only override sqref if it's missing or
  // tracked the previous activeCell, so explicit selections survive.
  const prevActive = selection.activeCell;
  if (selection.sqref === undefined || selection.sqref === prevActive) {
    selection.sqref = ref;
  }
  selection.activeCell = ref;
  view.selection = selection;
}

/**
 * Set the selected range (sqref) on the primary SheetView. Accepts
 * a single cell ("A1"), a single range ("A1:B5"), or a multi-cell
 * range string ("A1 C3:D4"). Leaves activeCell untouched unless
 * absent — in which case it's set to the first ref of `sqref`.
 */
export function setSelectedRange(ws: Worksheet, sqref: string): void {
  const view = ensurePrimaryView(ws);
  const selection = view.selection ?? {};
  selection.sqref = sqref;
  if (selection.activeCell === undefined) {
    const first = sqref.split(/\s+/)[0]?.split(':')[0];
    if (first) selection.activeCell = first;
  }
  view.selection = selection;
}

/**
 * Set values across a rectangular range from a 2-D array. `rows[0]`
 * is laid down starting at the top-left of `range`; subsequent rows
 * follow. `null` / `undefined` entries skip the cell. Useful for
 * dropping a header + data block in one call.
 */
export function setRangeValues(
  ws: Worksheet,
  range: string,
  rows: ReadonlyArray<ReadonlyArray<CellValue | null | undefined>>,
): void {
  const { minRow, minCol } = parseRange(range);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    for (let j = 0; j < row.length; j++) {
      const v = row[j];
      if (v === null || v === undefined) continue;
      setCell(ws, minRow + i, minCol + j, v);
    }
  }
}

/**
 * Iterate over every cell coordinate in a range, calling `visit`
 * once per (row, col). Allocates the cell on first touch so callers
 * can mutate it freely.
 */
export function applyToRange(
  ws: Worksheet,
  range: string,
  visit: (cell: Cell, row: number, col: number) => void,
): void {
  const { minRow, maxRow, minCol, maxCol } = parseRange(range);
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      let cell = ws.rows.get(r)?.get(c);
      if (!cell) cell = setCell(ws, r, c);
      visit(cell, r, c);
    }
  }
}

/**
 * Read a rectangular range as a dense 2-D array of values. Empty
 * cells yield `null`. The shape is `[maxRow - minRow + 1] ×
 * [maxCol - minCol + 1]`. Inverse of {@link setRangeValues}.
 */
export function getRangeValues(ws: Worksheet, range: string): (CellValue | null)[][] {
  const { minRow, maxRow, minCol, maxCol } = parseRange(range);
  const rowsOut: (CellValue | null)[][] = [];
  for (let r = minRow; r <= maxRow; r++) {
    const rowMap = ws.rows.get(r);
    const row: (CellValue | null)[] = [];
    for (let c = minCol; c <= maxCol; c++) {
      const cell = rowMap?.get(c);
      row.push(cell ? cell.value : null);
    }
    rowsOut.push(row);
  }
  return rowsOut;
}

/**
 * Read all populated values in a single column. Returns one `(CellValue
 * | null)` per row in `[minRow, maxRow]` (defaults to row 1 ..
 * `getMaxRow(ws)`). Empty cells yield `null`. Returns `[]` when the
 * worksheet is empty.
 */
export function getColumnValues(
  ws: Worksheet,
  col: number,
  opts: { minRow?: number; maxRow?: number } = {},
): (CellValue | null)[] {
  const max = getMaxRow(ws);
  if (max < 1) return [];
  const minRow = opts.minRow ?? 1;
  const maxRow = opts.maxRow ?? max;
  const out: (CellValue | null)[] = [];
  for (let r = minRow; r <= maxRow; r++) {
    const cell = ws.rows.get(r)?.get(col);
    out.push(cell ? cell.value : null);
  }
  return out;
}

/**
 * Read all populated values in a single row. Returns one `(CellValue
 * | null)` per column in `[minCol, maxCol]` (defaults to col 1 ..
 * `getMaxCol(ws)` when the row exists, otherwise `[]`).
 */
export function getRowValues(
  ws: Worksheet,
  row: number,
  opts: { minCol?: number; maxCol?: number } = {},
): (CellValue | null)[] {
  const rowMap = ws.rows.get(row);
  if (!rowMap || rowMap.size === 0) {
    if (opts.minCol === undefined && opts.maxCol === undefined) return [];
  }
  const minCol = opts.minCol ?? 1;
  const maxCol =
    opts.maxCol ??
    (rowMap ? Math.max(...rowMap.keys()) : 0);
  if (maxCol < minCol) return [];
  const out: (CellValue | null)[] = [];
  for (let c = minCol; c <= maxCol; c++) {
    const cell = rowMap?.get(c);
    out.push(cell ? cell.value : null);
  }
  return out;
}

// ---- column / row dimensions ----------------------------------------------

/**
 * Look up the ColumnDimension covering `col`. The search walks every
 * registered entry's `min..max` range; that's fine for the typical
 * spreadsheet (a handful of column entries) and stays simple.
 */
export function getColumnDimension(ws: Worksheet, col: number): ColumnDimension | undefined {
  for (const dim of ws.columnDimensions.values()) {
    if (col >= dim.min && col <= dim.max) return dim;
  }
  return undefined;
}

/**
 * Set a single-column ColumnDimension entry covering `col`. Shadows any
 * existing run that overlaps — runs are not split for now (callers that
 * need range-spanning entries can write directly into
 * `ws.columnDimensions`).
 */
export function setColumnDimension(
  ws: Worksheet,
  col: number,
  opts: Partial<Omit<ColumnDimension, 'min' | 'max'>>,
): ColumnDimension {
  validateRowCol(1, col);
  // Strip any existing entry that covers this column. Multi-col runs that
  // straddle `col` are dropped wholesale — phase-5 minimum scope.
  for (const [key, dim] of ws.columnDimensions) {
    if (col >= dim.min && col <= dim.max) ws.columnDimensions.delete(key);
  }
  const entry = makeColumnDimension(col, opts);
  ws.columnDimensions.set(col, entry);
  return entry;
}

/** Convenience: set a column's width, leaving other fields untouched. */
export function setColumnWidth(ws: Worksheet, col: number, width: number): ColumnDimension {
  const existing = getColumnDimension(ws, col);
  return setColumnDimension(ws, col, { ...existing, width, customWidth: true });
}

/** Convenience: hide a column. */
export function hideColumn(ws: Worksheet, col: number): ColumnDimension {
  const existing = getColumnDimension(ws, col);
  return setColumnDimension(ws, col, { ...existing, hidden: true });
}

/**
 * Set the default column width (characters) for cells without an
 * explicit ColumnDimension entry. Mirrors Excel's "Default Width"
 * dialog. Pass `undefined` to clear.
 */
export function setDefaultColumnWidth(ws: Worksheet, width: number | undefined): void {
  if (width === undefined) {
    delete (ws as { defaultColumnWidth?: number }).defaultColumnWidth;
    return;
  }
  if (!Number.isFinite(width) || width < 0) {
    throw new OpenXmlSchemaError(`setDefaultColumnWidth: width must be a non-negative number; got ${width}`);
  }
  ws.defaultColumnWidth = width;
}

/**
 * Set the default row height (points) for rows without an explicit
 * RowDimension entry. Mirrors Excel's "Default Row Height" dialog.
 * Pass `undefined` to clear.
 */
export function setDefaultRowHeight(ws: Worksheet, height: number | undefined): void {
  if (height === undefined) {
    delete (ws as { defaultRowHeight?: number }).defaultRowHeight;
    return;
  }
  if (!Number.isFinite(height) || height < 0) {
    throw new OpenXmlSchemaError(`setDefaultRowHeight: height must be a non-negative number; got ${height}`);
  }
  ws.defaultRowHeight = height;
}

/**
 * Approximate autofit for a column. Scans every populated cell in
 * `col` (or in `[opts.minRow, opts.maxRow]`), measures `cellValueAsString`
 * length, and sets the column width to `max(length) + padding`,
 * clamped to `[opts.min ?? 4, opts.max ?? 80]` and bounded above by
 * Excel's hard limit of 255.
 *
 * Note: this is a string-length approximation — Excel sizes columns
 * with the font's actual character-width metrics. For plain ASCII in
 * the default Calibri 11 face the result is usually within ±1 width
 * unit; CJK / wide glyphs need extra padding.
 */
export function autofitColumn(
  ws: Worksheet,
  col: number,
  opts: { minRow?: number; maxRow?: number; padding?: number; min?: number; max?: number } = {},
): ColumnDimension | undefined {
  const padding = opts.padding ?? 2;
  const minWidth = opts.min ?? 4;
  const maxWidth = Math.min(opts.max ?? 80, 255);
  const minRow = opts.minRow ?? 1;
  const maxRow = opts.maxRow ?? getMaxRow(ws);
  if (maxRow < minRow) return undefined;
  let widest = 0;
  for (let r = minRow; r <= maxRow; r++) {
    const cell = ws.rows.get(r)?.get(col);
    if (!cell) continue;
    const s = cellValueAsString(cell.value);
    if (s.length > widest) widest = s.length;
  }
  if (widest === 0) return undefined;
  const width = Math.max(minWidth, Math.min(maxWidth, widest + padding));
  return setColumnWidth(ws, col, width);
}

/**
 * Approximate autofit for every column with at least one populated
 * cell. Walks the worksheet once collecting per-column widest-length
 * + applies {@link autofitColumn} per column. `opts` flows through
 * unchanged.
 */
export function autofitColumns(
  ws: Worksheet,
  opts: { padding?: number; min?: number; max?: number } = {},
): void {
  const padding = opts.padding ?? 2;
  const minWidth = opts.min ?? 4;
  const maxWidth = Math.min(opts.max ?? 80, 255);
  const widest = new Map<number, number>();
  for (const rowMap of ws.rows.values()) {
    for (const [col, cell] of rowMap) {
      const s = cellValueAsString(cell.value);
      const cur = widest.get(col) ?? 0;
      if (s.length > cur) widest.set(col, s.length);
    }
  }
  for (const [col, w] of widest) {
    if (w === 0) continue;
    setColumnWidth(ws, col, Math.max(minWidth, Math.min(maxWidth, w + padding)));
  }
}

/**
 * Set widths for many columns in one call. `widths` maps either:
 * - an array `[12, 16, 20]` interpreted positionally starting at
 *   column `startCol` (default 1), or
 * - a `Record<number, number>` keyed by 1-based column index.
 * Each entry sets `customWidth: true`.
 */
export function setColumnWidths(
  ws: Worksheet,
  widths: ReadonlyArray<number> | Record<number, number>,
  startCol = 1,
): void {
  if (Array.isArray(widths)) {
    for (let i = 0; i < widths.length; i++) {
      const w = widths[i];
      if (typeof w !== 'number' || !Number.isFinite(w)) continue;
      setColumnWidth(ws, startCol + i, w);
    }
  } else {
    for (const [k, w] of Object.entries(widths as Record<number, number>)) {
      const col = Number.parseInt(k, 10);
      if (!Number.isInteger(col) || col < 1) continue;
      if (typeof w !== 'number' || !Number.isFinite(w)) continue;
      setColumnWidth(ws, col, w);
    }
  }
}

/** Look up a row's dimension entry. */
export function getRowDimension(ws: Worksheet, row: number): RowDimension | undefined {
  return ws.rowDimensions.get(row);
}

export function setRowDimension(ws: Worksheet, row: number, opts: Partial<RowDimension>): RowDimension {
  validateRowCol(row, 1);
  const entry = makeRowDimension(opts);
  ws.rowDimensions.set(row, entry);
  return entry;
}

/** Convenience: set a row's height, marking customHeight=true. */
export function setRowHeight(ws: Worksheet, row: number, height: number): RowDimension {
  const existing = getRowDimension(ws, row);
  return setRowDimension(ws, row, { ...existing, height, customHeight: true });
}

/**
 * Set heights for many rows in one call. `heights` accepts an array
 * (positional from `startRow`, default 1) or a `Record<number, number>`
 * keyed by 1-based row index. Each entry sets `customHeight: true`.
 */
export function setRowHeights(
  ws: Worksheet,
  heights: ReadonlyArray<number> | Record<number, number>,
  startRow = 1,
): void {
  if (Array.isArray(heights)) {
    for (let i = 0; i < heights.length; i++) {
      const h = heights[i];
      if (typeof h !== 'number' || !Number.isFinite(h)) continue;
      setRowHeight(ws, startRow + i, h);
    }
  } else {
    for (const [k, h] of Object.entries(heights as Record<number, number>)) {
      const row = Number.parseInt(k, 10);
      if (!Number.isInteger(row) || row < 1) continue;
      if (typeof h !== 'number' || !Number.isFinite(h)) continue;
      setRowHeight(ws, row, h);
    }
  }
}

/** Convenience: hide a row. */
export function hideRow(ws: Worksheet, row: number): RowDimension {
  const existing = getRowDimension(ws, row);
  return setRowDimension(ws, row, { ...existing, hidden: true });
}

// ---- hyperlinks -----------------------------------------------------------

/**
 * Replace any prior hyperlink on the same `ref` with the given options.
 * Pass `{ target }` for an external URL, `{ location }` for an internal
 * jump, or both. Returns the resulting Hyperlink record.
 */
export function setHyperlink(
  ws: Worksheet,
  ref: string,
  opts: { target?: string; location?: string; display?: string; tooltip?: string },
): Hyperlink {
  if (opts.target === undefined && opts.location === undefined) {
    throw new OpenXmlSchemaError('setHyperlink: one of target / location is required');
  }
  removeHyperlink(ws, ref);
  const hl = makeHyperlink({ ref, ...opts });
  ws.hyperlinks.push(hl);
  return hl;
}

/** Remove the hyperlink registered against `ref`. Returns true if anything was removed. */
export function removeHyperlink(ws: Worksheet, ref: string): boolean {
  const i = ws.hyperlinks.findIndex((h) => h.ref === ref);
  if (i < 0) return false;
  ws.hyperlinks.splice(i, 1);
  return true;
}

/** Look up a hyperlink by its ref. */
export function getHyperlink(ws: Worksheet, ref: string): Hyperlink | undefined {
  return ws.hyperlinks.find((h) => h.ref === ref);
}

/** Read-only snapshot of every hyperlink on the sheet. */
export function listHyperlinks(ws: Worksheet): ReadonlyArray<Hyperlink> {
  return ws.hyperlinks;
}

// ---- data validations ----------------------------------------------------

/** Append a DataValidation entry. */
export function addDataValidation(ws: Worksheet, dv: DataValidation): DataValidation {
  ws.dataValidations.push(dv);
  return dv;
}

/** Drop every validation whose sqref overlaps `ref` (string parse). Returns count removed. */
export function removeDataValidations(ws: Worksheet, predicate: (dv: DataValidation) => boolean): number {
  const before = ws.dataValidations.length;
  ws.dataValidations = ws.dataValidations.filter((dv) => !predicate(dv));
  return before - ws.dataValidations.length;
}

/** Read-only snapshot of every data validation block on the sheet. */
export function listDataValidations(ws: Worksheet): ReadonlyArray<DataValidation> {
  return ws.dataValidations;
}

// ---- autoFilter ----------------------------------------------------------

/** Set or replace the worksheet's AutoFilter. Pass `undefined` to clear. */
export function setAutoFilter(ws: Worksheet, filter: AutoFilter | undefined): void {
  if (filter === undefined) {
    delete ws.autoFilter;
    return;
  }
  ws.autoFilter = filter;
}

/** Read the current AutoFilter, if any. */
export function getAutoFilter(ws: Worksheet): AutoFilter | undefined {
  return ws.autoFilter;
}

// ---- tables --------------------------------------------------------------

/** Append a table. The id and displayName must be workbook-unique — the caller is responsible. */
export function addTable(ws: Worksheet, table: TableDefinition): TableDefinition {
  ws.tables.push(table);
  return table;
}

/** Look up a table by displayName. */
export function getTable(ws: Worksheet, displayName: string): TableDefinition | undefined {
  return ws.tables.find((t) => t.displayName === displayName);
}

/** Read-only snapshot of every Excel table defined on the sheet. */
export function listTables(ws: Worksheet): ReadonlyArray<TableDefinition> {
  return ws.tables;
}

/** Drop a table by displayName. Returns true when something was removed. */
export function removeTable(ws: Worksheet, displayName: string): boolean {
  const i = ws.tables.findIndex((t) => t.displayName === displayName);
  if (i < 0) return false;
  ws.tables.splice(i, 1);
  return true;
}

// ---- legacy comments -----------------------------------------------------

/** Add or replace the comment at `ref`. */
export function setComment(ws: Worksheet, opts: { ref: string; author: string; text: string }): LegacyComment {
  const i = ws.legacyComments.findIndex((c) => c.ref === opts.ref);
  const c = makeLegacyComment(opts);
  if (i < 0) ws.legacyComments.push(c);
  else ws.legacyComments[i] = c;
  return c;
}

export function getComment(ws: Worksheet, ref: string): LegacyComment | undefined {
  return ws.legacyComments.find((c) => c.ref === ref);
}

export function removeComment(ws: Worksheet, ref: string): boolean {
  const i = ws.legacyComments.findIndex((c) => c.ref === ref);
  if (i < 0) return false;
  ws.legacyComments.splice(i, 1);
  return true;
}

/** Read-only snapshot of every legacy comment on the sheet. */
export function listComments(ws: Worksheet): ReadonlyArray<LegacyComment> {
  return ws.legacyComments;
}

/**
 * Rename every comment authored by `oldName` to `newName`. Returns
 * the number of comments updated. Useful when consolidating comments
 * after a team handoff (Excel's commentsN.xml dedups authors at save
 * time, so a single rename collapses cleanly).
 */
export function renameCommentAuthor(ws: Worksheet, oldName: string, newName: string): number {
  let n = 0;
  for (let i = 0; i < ws.legacyComments.length; i++) {
    const c = ws.legacyComments[i];
    if (c && c.author === oldName) {
      ws.legacyComments[i] = makeLegacyComment({ ref: c.ref, author: newName, text: c.text });
      n++;
    }
  }
  return n;
}

/** Filter every legacy comment by author. */
export function findCommentsByAuthor(ws: Worksheet, author: string): ReadonlyArray<LegacyComment> {
  return ws.legacyComments.filter((c) => c.author === author);
}

// ---- conditional formatting ----------------------------------------------

/** Append a conditional formatting block. */
export function addConditionalFormatting(ws: Worksheet, cf: ConditionalFormatting): ConditionalFormatting {
  ws.conditionalFormatting.push(cf);
  return cf;
}

/** All conditional formatting blocks (read-only view). */
export function getConditionalFormatting(ws: Worksheet): ReadonlyArray<ConditionalFormatting> {
  return ws.conditionalFormatting;
}

// ---- cell watches / ignored errors --------------------------------------

/** Pin a cell to the Watch Window. Returns the pushed entry. */
export function addCellWatch(ws: Worksheet, watch: CellWatch): CellWatch {
  ws.cellWatches.push(watch);
  return watch;
}

/** Remove cell watches matching `predicate`. Returns the count removed. */
export function removeCellWatches(ws: Worksheet, predicate: (w: CellWatch) => boolean): number {
  const before = ws.cellWatches.length;
  ws.cellWatches = ws.cellWatches.filter((w) => !predicate(w));
  return before - ws.cellWatches.length;
}

/** Append an ignored-error region. */
export function addIgnoredError(ws: Worksheet, ie: IgnoredError): IgnoredError {
  ws.ignoredErrors.push(ie);
  return ie;
}

/** Remove ignored-error entries matching `predicate`. Returns the count removed. */
export function removeIgnoredErrors(ws: Worksheet, predicate: (ie: IgnoredError) => boolean): number {
  const before = ws.ignoredErrors.length;
  ws.ignoredErrors = ws.ignoredErrors.filter((ie) => !predicate(ie));
  return before - ws.ignoredErrors.length;
}
