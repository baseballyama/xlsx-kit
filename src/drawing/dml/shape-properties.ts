// DrawingML shape properties. Per docs/plan/08-charts-drawings.md §4.7.
//
// `<a:spPr>` is the universal "how should this be drawn" wrapper that
// every chart element (chartSpace, chart container, plotArea, series,
// dataPoint, axis, …) accepts. The model intentionally splits the
// wrapper (this file) from the leaves (colors / fill / line / geometry /
// effect / text) so the slot grows by attribute as new primitive
// modules land.

import type { EffectsRef } from './effect';
import type { Fill } from './fill';
import type { Geometry } from './geometry';
import type { LineProperties } from './line';

export type BlackWhiteMode =
  | 'clr'
  | 'auto'
  | 'gray'
  | 'ltGray'
  | 'invGray'
  | 'grayWhite'
  | 'blackGray'
  | 'blackWhite'
  | 'black'
  | 'white'
  | 'hidden';

export interface Point2D {
  x: number;
  y: number;
}

export interface PositiveSize2D {
  cx: number;
  cy: number;
}

/** `<a:xfrm>`. Position / size / rotation transformation. */
export interface Transform2D {
  off?: Point2D;
  ext?: PositiveSize2D;
  /** Rotation in 60_000-ths of a degree (0..21_600_000). */
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
  chOff?: Point2D;
  chExt?: PositiveSize2D;
}

/**
 * `<a:spPr>` wrapper. 3-D / text-body slots will be added as their
 * primitive modules land — kept absent here so the type stays closed
 * under the modules currently shipped.
 */
export interface ShapeProperties {
  bwMode?: BlackWhiteMode;
  xfrm?: Transform2D;
  geometry?: Geometry;
  fill?: Fill;
  ln?: LineProperties;
  /** Either `<a:effectLst>` (kind: 'lst') or `<a:effectDag>` (kind: 'dag'). */
  effects?: EffectsRef;
}

export const makeShapeProperties = (opts: Partial<ShapeProperties> = {}): ShapeProperties => ({
  ...(opts.bwMode !== undefined ? { bwMode: opts.bwMode } : {}),
  ...(opts.xfrm ? { xfrm: opts.xfrm } : {}),
  ...(opts.geometry ? { geometry: opts.geometry } : {}),
  ...(opts.fill ? { fill: opts.fill } : {}),
  ...(opts.ln ? { ln: opts.ln } : {}),
  ...(opts.effects ? { effects: opts.effects } : {}),
});
