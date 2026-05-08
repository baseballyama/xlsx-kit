// DifferentialStyle (DXF) — partial style overlay used by conditional
// formatting and Excel-table style elements. Mirrors
// openpyxl/openpyxl/styles/differential.py.
//
// Where a NamedStyle / CellXf carries a *complete* style snapshot, a
// DifferentialStyle carries only the components that override the base.
// The Stylesheet keeps DXFs in their own pool (`dxfs`) and conditional-
// formatting rules / table-style elements reference them by index.

import { stableStringify } from '../../utils/stable-stringify';
import type { Alignment } from './alignment';
import type { Border } from './borders';
import type { Fill } from './fills';
import type { Font } from './fonts';
import type { NumberFormat } from './numbers';
import type { Protection } from './protection';
import type { Stylesheet } from './stylesheet';

/**
 * Differential ("partial") style. Every component is optional; only
 * the set fields override the base style of whatever the DXF is
 * applied to.
 */
export interface DifferentialStyle {
  readonly font?: Font;
  readonly fill?: Fill;
  readonly border?: Border;
  readonly alignment?: Alignment;
  readonly protection?: Protection;
  /**
   * NumberFormat carries both the id and the format code so DXFs can
   * reference custom user-defined number formats without ambiguity.
   */
  readonly numFmt?: NumberFormat;
}

export function makeDifferentialStyle(opts: Partial<DifferentialStyle> = {}): DifferentialStyle {
  const out: { -readonly [K in keyof DifferentialStyle]: DifferentialStyle[K] } = {};
  if (opts.font !== undefined) out.font = opts.font;
  if (opts.fill !== undefined) out.fill = opts.fill;
  if (opts.border !== undefined) out.border = opts.border;
  if (opts.alignment !== undefined) out.alignment = opts.alignment;
  if (opts.protection !== undefined) out.protection = opts.protection;
  if (opts.numFmt !== undefined) out.numFmt = opts.numFmt;
  return Object.freeze(out);
}

/**
 * Stylesheet pool extension. The DXF list is allocated lazily on first
 * use so empty stylesheets stay slim.
 */
export interface StylesheetWithDxfs extends Stylesheet {
  dxfs?: DifferentialStyle[];
  _dxfIdByKey?: Map<string, number>;
}

/**
 * Add a DifferentialStyle to the stylesheet's `dxfs` pool. Returns
 * the 0-based index. Idempotent on structural equality (via
 * stableStringify).
 */
export function addDxf(ss: Stylesheet, dxf: DifferentialStyle): number {
  const w = ss as StylesheetWithDxfs;
  if (w.dxfs === undefined) w.dxfs = [];
  if (w._dxfIdByKey === undefined) w._dxfIdByKey = new Map();
  const key = stableStringify(dxf);
  const cached = w._dxfIdByKey.get(key);
  if (cached !== undefined) return cached;
  const id = w.dxfs.length;
  w.dxfs.push(dxf);
  w._dxfIdByKey.set(key, id);
  return id;
}

/** Read-only access to the dxfs pool. Returns `[]` when none registered. */
export function getDxfs(ss: Stylesheet): ReadonlyArray<DifferentialStyle> {
  return (ss as StylesheetWithDxfs).dxfs ?? [];
}
