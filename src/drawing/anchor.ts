// DrawingML anchors.
//
// An anchor positions a drawing (chart / image / shape) inside a worksheet.
// ECMA-376 §20.5.2 defines three kinds:
//
// absolute — fixed (x, y) and (cx, cy) in EMU; ignores cell layout. oneCell —
// pinned at `from` cell, fixed extent. Resizes with the
//               cell only on its top-left corner.
// twoCell — anchored to `from` and `to` cells. Resizes / moves with
//               both corners depending on `editAs`.
//
// Coordinates are EMU (English Metric Units). 1 inch = 914400 EMU, 1 cm =
// 360000 EMU, 1 px = 9525 EMU at 96 dpi.

import { columnIndexFromLetter } from '../utils/coordinate';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { emuFromPx } from '../utils/units';

/** EMU = English Metric Units. Drawing coordinates are stored in EMU on the wire. */
export interface Point2D {
  /** Horizontal offset in EMU. */
  x: number;
  /** Vertical offset in EMU. */
  y: number;
}

export interface PositiveSize2D {
  cx: number;
  cy: number;
}

/**
 * `from` / `to` corners reference a cell by 0-based column + row index plus an
 * EMU offset within the cell. Excel's wire format is 0-based here even though
 * cell references in formulas / sheetData are 1-based.
 */
export interface AnchorMarker {
  col: number;
  colOff: number;
  row: number;
  rowOff: number;
}

export type DrawingAnchor =
  | { kind: 'absolute'; pos: Point2D; ext: PositiveSize2D }
  | { kind: 'oneCell'; from: AnchorMarker; ext: PositiveSize2D }
  | { kind: 'twoCell'; from: AnchorMarker; to: AnchorMarker; editAs?: 'twoCell' | 'oneCell' | 'absolute' };

const A1_RE = /^([A-Za-z]{1,3})([1-9][0-9]*)$/;

/**
 * Convert a cell ref ("A1", "C5") to a 0-based AnchorMarker with `colOff =
 * rowOff = 0`. Throws on malformed refs.
 */
export function anchorMarkerFromCellRef(ref: string): AnchorMarker {
  const m = A1_RE.exec(ref);
  if (!m || m[1] === undefined || m[2] === undefined) {
    throw new OpenXmlSchemaError(`anchorMarkerFromCellRef: invalid coordinate "${ref}"`);
  }
  const col = columnIndexFromLetter(m[1]) - 1; // 0-based
  const row = Number.parseInt(m[2], 10) - 1; // 0-based
  return { col, colOff: 0, row, rowOff: 0 };
}

/** Build an absolute anchor from an (x, y, cx, cy) EMU bundle. */
export function makeAbsoluteAnchor(opts: { x: number; y: number; cx: number; cy: number }): DrawingAnchor {
  return { kind: 'absolute', pos: { x: opts.x, y: opts.y }, ext: { cx: opts.cx, cy: opts.cy } };
}

/** Build a one-cell anchor pinned at `from` with an explicit pixel extent. */
export function makeOneCellAnchor(opts: {
  from: string | AnchorMarker;
  widthPx: number;
  heightPx: number;
}): DrawingAnchor {
  const from = typeof opts.from === 'string' ? anchorMarkerFromCellRef(opts.from) : opts.from;
  return {
    kind: 'oneCell',
    from,
    ext: { cx: emuFromPx(opts.widthPx), cy: emuFromPx(opts.heightPx) },
  };
}

/**
 * Build a two-cell anchor from cell-ref pairs (or pre-built markers). Defaults
 * to `editAs='twoCell'` — drag both corners with the cells.
 */
export function makeTwoCellAnchor(opts: {
  from: string | AnchorMarker;
  to: string | AnchorMarker;
  editAs?: 'twoCell' | 'oneCell' | 'absolute';
}): DrawingAnchor {
  const from = typeof opts.from === 'string' ? anchorMarkerFromCellRef(opts.from) : opts.from;
  const to = typeof opts.to === 'string' ? anchorMarkerFromCellRef(opts.to) : opts.to;
  return {
    kind: 'twoCell',
    from,
    to,
    ...(opts.editAs !== undefined ? { editAs: opts.editAs } : {}),
  };
}
