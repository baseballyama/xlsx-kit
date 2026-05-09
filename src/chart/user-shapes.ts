// User shapes (chartDrawing) — annotations on a chart.
//
// `xl/drawings/chartDrawingN.xml` is referenced from a `<c:chartSpace>` via
// `<c:userShapes r:id="...">`. It contains text boxes, arrows, and other shapes
// positioned either relatively (0..1 of chart width / height) or absolutely
// (relative anchor + EMU extent).

import type { ShapeProperties, PositiveSize2D } from '../drawing/dml/shape-properties';
import type { TextBody } from '../drawing/dml/text';

/** `<cdr:from>` / `<cdr:to>` marker. Decimal in 0..1. */
export interface ChartRelativeMarker {
  x: number;
  y: number;
}

/** `<cdr:sp>` shape leaf (text box / arrow / preset shape with optional text). */
export interface ChartDrawingShape {
  /** cNvPr id (1-based, unique per chartDrawing). */
  id: number;
  name?: string;
  descr?: string;
  hidden?: boolean;
  /** When true, emit `<cdr:nvSpPr><cdr:cNvSpPr txBox="1"/>` so Excel renders the shape as a text-box. */
  txBox?: boolean;
  spPr?: ShapeProperties;
  txBody?: TextBody;
}

/** `<cdr:pic>` picture leaf. */
export interface ChartDrawingPicture {
  id: number;
  name?: string;
  descr?: string;
  /** rels-resolved id used by `<a:blip r:embed>`. */
  embedRId?: string;
  spPr?: ShapeProperties;
}

export type UserShapeContent =
  | { kind: 'shape'; shape: ChartDrawingShape }
  | { kind: 'picture'; picture: ChartDrawingPicture };

export type UserShapeAnchor =
  | {
      kind: 'relSize';
      from: ChartRelativeMarker;
      to: ChartRelativeMarker;
      content: UserShapeContent;
    }
  | {
      kind: 'absSize';
      from: ChartRelativeMarker;
      ext: PositiveSize2D;
      content: UserShapeContent;
    };

export interface ChartDrawing {
  shapes: UserShapeAnchor[];
}

export const makeChartDrawing = (shapes: UserShapeAnchor[] = []): ChartDrawing => ({ shapes });

export const makeChartShape = (opts: {
  id: number;
  name?: string;
  descr?: string;
  hidden?: boolean;
  txBox?: boolean;
  spPr?: ShapeProperties;
  txBody?: TextBody;
}): ChartDrawingShape => ({
  id: opts.id,
  ...(opts.name !== undefined ? { name: opts.name } : {}),
  ...(opts.descr !== undefined ? { descr: opts.descr } : {}),
  ...(opts.hidden !== undefined ? { hidden: opts.hidden } : {}),
  ...(opts.txBox !== undefined ? { txBox: opts.txBox } : {}),
  ...(opts.spPr ? { spPr: opts.spPr } : {}),
  ...(opts.txBody ? { txBody: opts.txBody } : {}),
});

export const makeRelSizeAnchor = (
  from: ChartRelativeMarker,
  to: ChartRelativeMarker,
  content: UserShapeContent,
): UserShapeAnchor => ({ kind: 'relSize', from, to, content });

export const makeAbsSizeAnchor = (
  from: ChartRelativeMarker,
  ext: PositiveSize2D,
  content: UserShapeContent,
): UserShapeAnchor => ({ kind: 'absSize', from, ext, content });
