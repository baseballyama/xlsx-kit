// DrawingML line properties.

import type { Fill } from './fill';

export type PresetDash =
  | 'solid'
  | 'dot'
  | 'dash'
  | 'lgDash'
  | 'dashDot'
  | 'lgDashDot'
  | 'lgDashDotDot'
  | 'sysDash'
  | 'sysDot'
  | 'sysDashDot'
  | 'sysDashDotDot';

export type LineDash = { kind: 'preset'; val: PresetDash } | { kind: 'custDash'; pattern: number[] };

export type LineCap = 'rnd' | 'sq' | 'flat';
export type LineCompound = 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';
export type LineAlign = 'ctr' | 'in';
export type LineJoin = { kind: 'round' } | { kind: 'bevel' } | { kind: 'miter'; lim?: number };

export type LineEndType = 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
export type LineEndSize = 'sm' | 'med' | 'lg';

export interface LineEnd {
  type?: LineEndType;
  w?: LineEndSize;
  len?: LineEndSize;
}

export interface LineProperties {
  /** Line width in EMU (1pt ≈ 12700 EMU). */
  w?: number;
  cap?: LineCap;
  cmpd?: LineCompound;
  algn?: LineAlign;
  fill?: Fill;
  dash?: LineDash;
  join?: LineJoin;
  headEnd?: LineEnd;
  tailEnd?: LineEnd;
}

export const makeLine = (opts: Partial<LineProperties> = {}): LineProperties => ({
  ...(opts.w !== undefined ? { w: opts.w } : {}),
  ...(opts.cap !== undefined ? { cap: opts.cap } : {}),
  ...(opts.cmpd !== undefined ? { cmpd: opts.cmpd } : {}),
  ...(opts.algn !== undefined ? { algn: opts.algn } : {}),
  ...(opts.fill ? { fill: opts.fill } : {}),
  ...(opts.dash ? { dash: opts.dash } : {}),
  ...(opts.join ? { join: opts.join } : {}),
  ...(opts.headEnd ? { headEnd: opts.headEnd } : {}),
  ...(opts.tailEnd ? { tailEnd: opts.tailEnd } : {}),
});
