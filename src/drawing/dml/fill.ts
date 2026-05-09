// DrawingML fills.

import type { DmlColor, DmlColorWithMods } from './colors';

/** Relative rectangle (0..100000-thousandths). */
export interface RelativeRect {
  l?: number;
  t?: number;
  r?: number;
  b?: number;
}

export interface GradientStop {
  /** Position along the gradient (0..100000). */
  pos: number;
  color: DmlColorWithMods;
}

export type GradientLineDir =
  | { kind: 'lin'; ang: number; scaled?: boolean }
  | { kind: 'path'; pathType: 'shape' | 'circle' | 'rect'; tileRect?: RelativeRect };

export type TileFlip = 'x' | 'y' | 'xy' | 'none';

export interface TileFill {
  tx?: number;
  ty?: number;
  sx?: number;
  sy?: number;
  flip?: TileFlip;
  algn?: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
}

/** Blip-source effects. The full ECMA-376 list is preserved here; `kind` discriminates. */
export type BlipEffect =
  | { kind: 'biLevel'; thresh: number }
  | { kind: 'blur'; rad: number; grow?: boolean }
  | { kind: 'clrChange'; useA?: boolean; clrFrom: DmlColor; clrTo: DmlColor }
  | { kind: 'clrRepl'; color: DmlColor }
  | { kind: 'duotone'; colors: [DmlColor, DmlColor] }
  | { kind: 'grayscl' }
  | { kind: 'lum'; bright?: number; contrast?: number }
  | { kind: 'tint'; hue?: number; amt?: number }
  | { kind: 'alphaModFix'; amt: number };

export interface Blip {
  embedRId?: string;
  linkRId?: string;
  cstate?: 'email' | 'screen' | 'print' | 'hqprint';
  effects?: BlipEffect[];
}

/** Discriminated union for any fill that can attach to a shape. */
export type Fill =
  | { kind: 'noFill' }
  | { kind: 'solidFill'; color: DmlColorWithMods }
  | {
      kind: 'gradFill';
      flip?: TileFlip;
      rotWithShape?: boolean;
      stops: GradientStop[];
      lineDir?: GradientLineDir;
    }
  | {
      kind: 'blipFill';
      blip: Blip;
      tile?: TileFill;
      stretch?: { fillRect?: RelativeRect };
      srcRect?: RelativeRect;
      dpi?: number;
      rotWithShape?: boolean;
    }
  | {
      kind: 'pattFill';
      preset: string;
      fgClr?: DmlColorWithMods;
      bgClr?: DmlColorWithMods;
    }
  | { kind: 'grpFill' };

export const makeNoFill = (): Fill => ({ kind: 'noFill' });
export const makeSolidFill = (color: DmlColorWithMods): Fill => ({ kind: 'solidFill', color });
export const makeGroupFill = (): Fill => ({ kind: 'grpFill' });

export const makeGradientFill = (opts: {
  stops: GradientStop[];
  flip?: TileFlip;
  rotWithShape?: boolean;
  lineDir?: GradientLineDir;
}): Fill => ({
  kind: 'gradFill',
  stops: opts.stops,
  ...(opts.flip !== undefined ? { flip: opts.flip } : {}),
  ...(opts.rotWithShape !== undefined ? { rotWithShape: opts.rotWithShape } : {}),
  ...(opts.lineDir ? { lineDir: opts.lineDir } : {}),
});

export const makePatternFill = (opts: {
  preset: string;
  fgClr?: DmlColorWithMods;
  bgClr?: DmlColorWithMods;
}): Fill => ({
  kind: 'pattFill',
  preset: opts.preset,
  ...(opts.fgClr ? { fgClr: opts.fgClr } : {}),
  ...(opts.bgClr ? { bgClr: opts.bgClr } : {}),
});

/** Excel's preset pattern names (54 entries, ECMA-376 §20.1.10.50; matches openpyxl). */
export const PRESET_PATTERN_NAMES: ReadonlyArray<string> = [
  'pct5',
  'pct10',
  'pct20',
  'pct25',
  'pct30',
  'pct40',
  'pct50',
  'pct60',
  'pct70',
  'pct75',
  'pct80',
  'pct90',
  'horz',
  'vert',
  'ltHorz',
  'ltVert',
  'dkHorz',
  'dkVert',
  'narHorz',
  'narVert',
  'dashHorz',
  'dashVert',
  'cross',
  'dnDiag',
  'upDiag',
  'ltDnDiag',
  'ltUpDiag',
  'dkDnDiag',
  'dkUpDiag',
  'wdDnDiag',
  'wdUpDiag',
  'dashDnDiag',
  'dashUpDiag',
  'diagCross',
  'smCheck',
  'lgCheck',
  'smGrid',
  'lgGrid',
  'dotGrid',
  'smConfetti',
  'lgConfetti',
  'horzBrick',
  'diagBrick',
  'solidDmnd',
  'openDmnd',
  'dotDmnd',
  'plaid',
  'sphere',
  'weave',
  'divot',
  'shingle',
  'wave',
  'trellis',
  'zigZag',
];
