// Cell ↔ Stylesheet bridge. Per docs/plan/04-core-model.md §3.6.
//
// A cell's `styleId` is an index into `Workbook.styles.cellXfs`. Each
// CellXf points at slots in the font / fill / border / numFmt pools.
// To "apply" a Font to a cell we therefore:
//   1. read the cell's current CellXf (or `defaultCellXf` if it points
//      at a slot that hasn't been allocated yet — common right after
//      `makeCell` since `cellXfs` starts empty)
//   2. resolve / register the new component in its pool
//   3. build a new CellXf that carries the new id + the matching
//      `apply*` flag (Excel needs the flag to honour the override over
//      the underlying NamedStyle)
//   4. dedup that CellXf via `addCellXf` and write the returned index
//      back to `c.styleId`
//
// Following the no-classes rule, this module is just a flat list of
// free functions; the workbook is passed in so callers don't need to
// thread the stylesheet manually.

import type { Cell } from '../cell/cell';
import { OpenXmlSchemaError } from '../utils/exceptions';
import type { Workbook } from '../workbook/workbook';
import { parseRange } from '../worksheet/cell-range';
import { setCell, type Worksheet } from '../worksheet/worksheet';
import type { Alignment } from './alignment';
import type { Border, SideStyle } from './borders';
import { DEFAULT_BORDER, makeBorder, makeSide } from './borders';
import type { Color } from './colors';
import { makeColor } from './colors';
import type { Fill } from './fills';
import { DEFAULT_EMPTY_FILL, makePatternFill } from './fills';
import type { Font } from './fonts';
import { DEFAULT_FONT, makeFont } from './fonts';
import { ensureBuiltinStyle } from './named-styles';
import { builtinFormatCode } from './numbers';
import type { Protection } from './protection';
import { DEFAULT_PROTECTION } from './protection';
import {
  addBorder,
  addCellXf,
  addFill,
  addFont,
  addNumFmt,
  type CellXf,
  defaultCellXf,
  type Stylesheet,
} from './stylesheet';

/** Default General number format code (numFmtId 0). */
const GENERAL_FORMAT_CODE = 'General';

/** Resolve a cell's current CellXf, falling back to defaults when unset. */
function currentXf(ss: Stylesheet, c: Cell): CellXf {
  return ss.cellXfs[c.styleId] ?? defaultCellXf();
}

// ---- read accessors --------------------------------------------------------

export function getCellFont(wb: Workbook, c: Cell): Font {
  const xf = currentXf(wb.styles, c);
  return wb.styles.fonts[xf.fontId] ?? DEFAULT_FONT;
}

export function getCellFill(wb: Workbook, c: Cell): Fill {
  const xf = currentXf(wb.styles, c);
  return wb.styles.fills[xf.fillId] ?? DEFAULT_EMPTY_FILL;
}

export function getCellBorder(wb: Workbook, c: Cell): Border {
  const xf = currentXf(wb.styles, c);
  return wb.styles.borders[xf.borderId] ?? DEFAULT_BORDER;
}

export function getCellAlignment(wb: Workbook, c: Cell): Alignment {
  return currentXf(wb.styles, c).alignment ?? {};
}

export function getCellProtection(wb: Workbook, c: Cell): Protection {
  return currentXf(wb.styles, c).protection ?? DEFAULT_PROTECTION;
}

/**
 * Returns the cell's number-format **code** (e.g. `"0.00"`, `"General"`).
 * Built-in IDs resolve through `builtinFormatCode`; custom IDs come from
 * the workbook's numFmts map.
 */
export function getCellNumberFormat(wb: Workbook, c: Cell): string {
  const id = currentXf(wb.styles, c).numFmtId;
  const builtin = builtinFormatCode(id);
  if (builtin !== undefined) return builtin;
  return wb.styles.numFmts.get(id) ?? GENERAL_FORMAT_CODE;
}

// ---- write accessors -------------------------------------------------------

/**
 * Reserve cellXfs[0] for the implicit default xf when the pool is
 * empty. Excel's `<c>` elements without an `s=` attribute resolve to
 * `cellXfs[0]`, so the first time a caller styles any cell we need
 * to make sure that slot stays the default — otherwise unstyled
 * cells in the same sheet would inherit the freshly added styled xf.
 *
 * Idempotent: calling this on a non-empty pool is a no-op.
 */
const reserveDefaultXfSlot = (wb: Workbook): void => {
  if (wb.styles.cellXfs.length === 0) addCellXf(wb.styles, defaultCellXf());
};

/**
 * Replace one field on the cell's CellXf. Centralises the dedup +
 * styleId update so each `setCell*` is a single dispatch.
 */
function applyXfPatch(wb: Workbook, c: Cell, patch: Partial<CellXf>): void {
  reserveDefaultXfSlot(wb);
  const next: CellXf = { ...currentXf(wb.styles, c), ...patch };
  c.styleId = addCellXf(wb.styles, next);
}

export function setCellFont(wb: Workbook, c: Cell, font: Font): void {
  const fontId = addFont(wb.styles, font);
  applyXfPatch(wb, c, { fontId, applyFont: true });
}

export function setCellFill(wb: Workbook, c: Cell, fill: Fill): void {
  const fillId = addFill(wb.styles, fill);
  applyXfPatch(wb, c, { fillId, applyFill: true });
}

export function setCellBorder(wb: Workbook, c: Cell, border: Border): void {
  const borderId = addBorder(wb.styles, border);
  applyXfPatch(wb, c, { borderId, applyBorder: true });
}

export function setCellAlignment(wb: Workbook, c: Cell, alignment: Alignment): void {
  applyXfPatch(wb, c, { alignment, applyAlignment: true });
}

export function setCellProtection(wb: Workbook, c: Cell, protection: Protection): void {
  applyXfPatch(wb, c, { protection, applyProtection: true });
}

/**
 * Set the cell's number format by its **code** string.
 * Built-in codes resolve to their canonical id; custom codes are
 * registered via `addNumFmt`.
 */
export function setCellNumberFormat(wb: Workbook, c: Cell, formatCode: string): void {
  const numFmtId = addNumFmt(wb.styles, formatCode);
  applyXfPatch(wb, c, { numFmtId, applyNumberFormat: true });
}

/**
 * Build a single CellXf id from a multi-axis style spec, then apply
 * it to every cell in `range`. The xf is registered once per style
 * shape, so a 1000-cell range allocates one xf — much faster than
 * looping `setCellStyle` per cell.
 */
export function setRangeStyle(
  wb: Workbook,
  ws: Worksheet,
  range: string,
  opts: {
    font?: Font;
    fill?: Fill;
    border?: Border;
    alignment?: Alignment;
    protection?: Protection;
    numberFormat?: string;
  },
): void {
  const patch: { -readonly [K in keyof CellXf]?: CellXf[K] } = {};
  if (opts.font !== undefined) {
    patch.fontId = addFont(wb.styles, opts.font);
    patch.applyFont = true;
  }
  if (opts.fill !== undefined) {
    patch.fillId = addFill(wb.styles, opts.fill);
    patch.applyFill = true;
  }
  if (opts.border !== undefined) {
    patch.borderId = addBorder(wb.styles, opts.border);
    patch.applyBorder = true;
  }
  if (opts.alignment !== undefined) {
    patch.alignment = opts.alignment;
    patch.applyAlignment = true;
  }
  if (opts.protection !== undefined) {
    patch.protection = opts.protection;
    patch.applyProtection = true;
  }
  if (opts.numberFormat !== undefined) {
    patch.numFmtId = addNumFmt(wb.styles, opts.numberFormat);
    patch.applyNumberFormat = true;
  }
  if (Object.keys(patch).length === 0) return;
  reserveDefaultXfSlot(wb);

  const { minRow, maxRow, minCol, maxCol } = parseRange(range);
  // Pre-register the xf for each existing cell — Excel dedupes by
  // value, so the inner xf-pool ends up the same shape for cells
  // already carrying part of the patch as for blanks.
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      let cell = ws.rows.get(r)?.get(c);
      if (!cell) cell = setCell(ws, r, c);
      const next: CellXf = { ...currentXf(wb.styles, cell), ...patch };
      cell.styleId = addCellXf(wb.styles, next);
    }
  }
}

/**
 * Combined cell-style setter. Each axis is independent — pass any
 * subset and the corresponding `applyXxx` flags get set on the
 * underlying CellXf. Avoids 5+ separate stylesheet round-trips when
 * a caller wants to style a single cell across multiple axes (Excel
 * dedupes the resulting xf record on every call).
 */
export function setCellStyle(
  wb: Workbook,
  c: Cell,
  opts: {
    font?: Font;
    fill?: Fill;
    border?: Border;
    alignment?: Alignment;
    protection?: Protection;
    numberFormat?: string;
  },
): void {
  const patch: { -readonly [K in keyof CellXf]?: CellXf[K] } = {};
  if (opts.font !== undefined) {
    patch.fontId = addFont(wb.styles, opts.font);
    patch.applyFont = true;
  }
  if (opts.fill !== undefined) {
    patch.fillId = addFill(wb.styles, opts.fill);
    patch.applyFill = true;
  }
  if (opts.border !== undefined) {
    patch.borderId = addBorder(wb.styles, opts.border);
    patch.applyBorder = true;
  }
  if (opts.alignment !== undefined) {
    patch.alignment = opts.alignment;
    patch.applyAlignment = true;
  }
  if (opts.protection !== undefined) {
    patch.protection = opts.protection;
    patch.applyProtection = true;
  }
  if (opts.numberFormat !== undefined) {
    patch.numFmtId = addNumFmt(wb.styles, opts.numberFormat);
    patch.applyNumberFormat = true;
  }
  if (Object.keys(patch).length === 0) return;
  applyXfPatch(wb, c, patch as Partial<CellXf>);
}

// ---- format presets -----------------------------------------------------
//
// Mirror Excel's "Format Cells → Number → Category" panel. Each preset
// builds the exact format-code Excel ships with, then routes through
// the existing `setCellNumberFormat` so the dedup pool sees the same
// string every time.

/**
 * Format a cell as currency. Produces one of:
 * - default → `"$#,##0.00"` (US dollar, 2 decimals)
 * - `{ symbol: "€" }` → `"€#,##0.00"`
 * - `{ symbol: "¥", decimals: 0 }` → `"¥#,##0"`
 * - `{ accounting: true }` → `"_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-"`
 *   (Excel's "Accounting" subtype with right-aligned symbol).
 */
export function setCellAsCurrency(
  wb: Workbook,
  c: Cell,
  opts: { symbol?: string; decimals?: number; accounting?: boolean } = {},
): void {
  const symbol = opts.symbol ?? '$';
  const decimals = opts.decimals ?? 2;
  const decTail = decimals > 0 ? `.${'0'.repeat(decimals)}` : '';
  const code = opts.accounting
    ? `_-${symbol}* #,##0${decTail}_-;-${symbol}* #,##0${decTail}_-;_-${symbol}* "-"${'?'.repeat(decimals)}_-;_-@_-`
    : `${symbol}#,##0${decTail}`;
  setCellNumberFormat(wb, c, code);
}

/**
 * Format a cell as a percentage. `decimals` defaults to 0 → `"0%"`;
 * `decimals: 2` → `"0.00%"`. The cell value is multiplied by 100 by
 * Excel during display.
 */
export function setCellAsPercent(wb: Workbook, c: Cell, decimals = 0): void {
  if (!Number.isInteger(decimals) || decimals < 0) {
    throw new OpenXmlSchemaError(`setCellAsPercent: decimals must be a non-negative integer; got ${decimals}`);
  }
  const code = decimals === 0 ? '0%' : `0.${'0'.repeat(decimals)}%`;
  setCellNumberFormat(wb, c, code);
}

/**
 * Format a cell as a date. `format` defaults to Excel's default
 * locale-independent ISO-style date `"yyyy-mm-dd"`. Common alternatives:
 * `"m/d/yyyy"`, `"dd-mmm-yy"`, `"yyyy-mm-dd hh:mm:ss"`.
 */
export function setCellAsDate(wb: Workbook, c: Cell, format = 'yyyy-mm-dd'): void {
  setCellNumberFormat(wb, c, format);
}

/**
 * Format a cell as a thousands-separated number. `decimals` defaults
 * to 0 → `"#,##0"`; `decimals: 2` → `"#,##0.00"`.
 */
export function setCellAsNumber(wb: Workbook, c: Cell, decimals = 0): void {
  if (!Number.isInteger(decimals) || decimals < 0) {
    throw new OpenXmlSchemaError(`setCellAsNumber: decimals must be a non-negative integer; got ${decimals}`);
  }
  const code = decimals === 0 ? '#,##0' : `#,##0.${'0'.repeat(decimals)}`;
  setCellNumberFormat(wb, c, code);
}

// ---- table-header preset ------------------------------------------------

/**
 * Apply Excel's stock "table header" formatting to a range: bold
 * white text on a dark fill, plus a thick bottom border. Override
 * any axis via `opts` — pass `bold: false` to drop the bold, or
 * `fillColor: 'FF305496'` for a different shade. Defaults match
 * Excel's "Table Style Medium 2" header row.
 */
export function formatAsHeader(
  wb: Workbook,
  ws: Worksheet,
  range: string,
  opts: {
    fillColor?: string | Partial<Color>;
    fontColor?: string | Partial<Color>;
    bold?: boolean;
    bottomBorder?: SideStyle | false;
    bottomBorderColor?: string | Partial<Color>;
  } = {},
): void {
  const fillColor = opts.fillColor === undefined ? 'FF305496' : typeof opts.fillColor === 'string' ? makeColor({ rgb: opts.fillColor }) : makeColor(opts.fillColor);
  const fontColor = opts.fontColor === undefined ? 'FFFFFFFF' : typeof opts.fontColor === 'string' ? makeColor({ rgb: opts.fontColor }) : makeColor(opts.fontColor);
  const bold = opts.bold ?? true;
  const borderStyle = opts.bottomBorder ?? 'medium';
  const borderColorObj = opts.bottomBorderColor === undefined ? undefined : typeof opts.bottomBorderColor === 'string' ? makeColor({ rgb: opts.bottomBorderColor }) : makeColor(opts.bottomBorderColor);

  const fillColorObj = typeof fillColor === 'string' ? makeColor({ rgb: fillColor }) : fillColor;
  const fontColorObj = typeof fontColor === 'string' ? makeColor({ rgb: fontColor }) : fontColor;

  const styleOpts: Parameters<typeof setRangeStyle>[3] = {
    font: makeFont({ bold, color: fontColorObj }),
    fill: makePatternFill({ patternType: 'solid', fgColor: fillColorObj }),
  };
  if (borderStyle !== false) {
    const side = makeSide({ style: borderStyle, ...(borderColorObj ? { color: borderColorObj } : {}) });
    styleOpts.border = makeBorder({ bottom: side });
  }
  setRangeStyle(wb, ws, range, styleOpts);
}

// ---- named / built-in style application --------------------------------

/**
 * Apply a built-in Excel style ("Heading 1" / "Total" / "Good" /
 * "Bad" / "Calculation" / etc.) to a single cell. Registers the
 * built-in on the Stylesheet (idempotent) and points the cell's xf
 * at it via `xfId` while inheriting the matching font/fill/border/
 * numFmt ids so the cell renders correctly on its own.
 *
 * Throws when `name` isn't in {@link BUILTIN_NAMED_STYLES}; use
 * {@link applyNamedStyle} for user-registered styles.
 */
export function applyBuiltinStyle(wb: Workbook, c: Cell, name: string): void {
  const xfId = ensureBuiltinStyle(wb.styles, name);
  applyNamedStyleByXfId(wb, c, xfId);
}

/**
 * Apply a NamedStyle that's already registered on the workbook (via
 * `addNamedStyle` or `ensureBuiltinStyle`) to a single cell, by name.
 */
export function applyNamedStyle(wb: Workbook, c: Cell, name: string): void {
  const entry = wb.styles._namedStyleByName?.get(name);
  if (entry === undefined) {
    throw new OpenXmlSchemaError(`applyNamedStyle: no named style "${name}" registered`);
  }
  applyNamedStyleByXfId(wb, c, entry.xfId);
}

const applyNamedStyleByXfId = (wb: Workbook, c: Cell, xfId: number): void => {
  const styleXf = wb.styles.cellStyleXfs[xfId];
  if (!styleXf) {
    throw new OpenXmlSchemaError(`applyNamedStyle: cellStyleXfs[${xfId}] missing`);
  }
  const patch: { -readonly [K in keyof CellXf]?: CellXf[K] } = {
    xfId,
    fontId: styleXf.fontId,
    fillId: styleXf.fillId,
    borderId: styleXf.borderId,
    numFmtId: styleXf.numFmtId,
  };
  if (styleXf.applyFont) patch.applyFont = true;
  if (styleXf.applyFill) patch.applyFill = true;
  if (styleXf.applyBorder) patch.applyBorder = true;
  if (styleXf.applyNumberFormat) patch.applyNumberFormat = true;
  if (styleXf.alignment !== undefined) {
    patch.alignment = styleXf.alignment;
    patch.applyAlignment = true;
  }
  if (styleXf.protection !== undefined) {
    patch.protection = styleXf.protection;
    patch.applyProtection = true;
  }
  applyXfPatch(wb, c, patch as Partial<CellXf>);
};

// ---- border presets ----------------------------------------------------

/**
 * Apply the same {@link SideStyle} to all four edges of a single cell.
 * Optional color via hex string or `Color` partial. Equivalent to
 * `setCellBorder(wb, c, makeBorder({ left, right, top, bottom: side }))`
 * with all four sides identical.
 */
export function setCellBorderAll(
  wb: Workbook,
  c: Cell,
  opts: { style: SideStyle; color?: string | Partial<Color> } = { style: 'thin' },
): void {
  const colorObj = opts.color === undefined ? undefined : typeof opts.color === 'string' ? makeColor({ rgb: opts.color }) : makeColor(opts.color);
  const side = makeSide({ style: opts.style, ...(colorObj ? { color: colorObj } : {}) });
  setCellBorder(wb, c, makeBorder({ left: side, right: side, top: side, bottom: side }));
}

/**
 * Draw an outer border around a rectangular range. Cells on the
 * perimeter receive a partial border (only the edges that face
 * outside the range); inner cells are unaffected unless `inner` is
 * provided, in which case every cell in the range receives a border
 * combining its perimeter edges with the `inner` style for the inside
 * edges.
 */
export function setRangeBorderBox(
  wb: Workbook,
  ws: Worksheet,
  range: string,
  opts: { style: SideStyle; color?: string | Partial<Color>; inner?: SideStyle } = { style: 'thin' },
): void {
  const { minRow, maxRow, minCol, maxCol } = parseRange(range);
  const colorObj = opts.color === undefined ? undefined : typeof opts.color === 'string' ? makeColor({ rgb: opts.color }) : makeColor(opts.color);
  const outer = makeSide({ style: opts.style, ...(colorObj ? { color: colorObj } : {}) });
  const inner =
    opts.inner !== undefined
      ? makeSide({ style: opts.inner, ...(colorObj ? { color: colorObj } : {}) })
      : undefined;
  for (let r = minRow; r <= maxRow; r++) {
    for (let col = minCol; col <= maxCol; col++) {
      const onTop = r === minRow;
      const onBottom = r === maxRow;
      const onLeft = col === minCol;
      const onRight = col === maxCol;
      // Skip cells that wouldn't get any styling (interior cells when no inner).
      if (!inner && !onTop && !onBottom && !onLeft && !onRight) continue;
      const sides: { -readonly [K in keyof Border]?: Border[K] } = {};
      const top = onTop ? outer : inner;
      const bottom = onBottom ? outer : inner;
      const left = onLeft ? outer : inner;
      const right = onRight ? outer : inner;
      if (top !== undefined) sides.top = top;
      if (bottom !== undefined) sides.bottom = bottom;
      if (left !== undefined) sides.left = left;
      if (right !== undefined) sides.right = right;
      let cell = ws.rows.get(r)?.get(col);
      if (!cell) cell = setCell(ws, r, col);
      setCellBorder(wb, cell, makeBorder(sides));
    }
  }
}
