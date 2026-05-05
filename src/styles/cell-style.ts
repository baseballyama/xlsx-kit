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
import type { Workbook } from '../workbook/workbook';
import { parseRange } from '../worksheet/cell-range';
import { setCell, type Worksheet } from '../worksheet/worksheet';
import type { Alignment } from './alignment';
import type { Border } from './borders';
import { DEFAULT_BORDER } from './borders';
import type { Fill } from './fills';
import { DEFAULT_EMPTY_FILL } from './fills';
import type { Font } from './fonts';
import { DEFAULT_FONT } from './fonts';
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
 * Replace one field on the cell's CellXf. Centralises the dedup +
 * styleId update so each `setCell*` is a single dispatch.
 */
function applyXfPatch(wb: Workbook, c: Cell, patch: Partial<CellXf>): void {
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
