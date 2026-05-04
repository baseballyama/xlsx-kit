// Border / Side value objects. Mirrors openpyxl/openpyxl/styles/borders.py.
//
// A `Border` describes the edges drawn around a cell. Each edge is a
// `Side` carrying a stroke style and an optional `Color`. Per
// docs/plan/04-core-model.md §3.1 these are plain readonly objects;
// the `make*` constructors freeze their results so the Stylesheet
// pool can dedupe by reference identity once we wire it up.

import { OpenXmlSchemaError } from '../utils/exceptions';
import { type Color, makeColor } from './colors';

export type SideStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'double'
  | 'hair'
  | 'dotted'
  | 'dashed'
  | 'dashDot'
  | 'dashDotDot'
  | 'mediumDashed'
  | 'mediumDashDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export const SIDE_STYLES: ReadonlyArray<SideStyle> = Object.freeze([
  'thin',
  'medium',
  'thick',
  'double',
  'hair',
  'dotted',
  'dashed',
  'dashDot',
  'dashDotDot',
  'mediumDashed',
  'mediumDashDot',
  'mediumDashDotDot',
  'slantDashDot',
]);

const SIDE_STYLE_SET: ReadonlySet<string> = new Set(SIDE_STYLES);

export interface Side {
  readonly style?: SideStyle;
  readonly color?: Color;
}

/** Build an immutable {@link Side}. */
export function makeSide(opts: Partial<Side> = {}): Side {
  const out: { -readonly [K in keyof Side]: Side[K] } = {};
  if (opts.style !== undefined) {
    if (!SIDE_STYLE_SET.has(opts.style)) {
      throw new OpenXmlSchemaError(`Side style must be one of [${SIDE_STYLES.join(', ')}]; got "${opts.style}"`);
    }
    out.style = opts.style;
  }
  if (opts.color !== undefined) {
    // Funnel through makeColor so we re-use its validation + freezing.
    out.color = Object.isFrozen(opts.color) ? opts.color : makeColor(opts.color);
  }
  return Object.freeze(out);
}

export interface Border {
  /** Left edge. */
  readonly left?: Side;
  readonly right?: Side;
  readonly top?: Side;
  readonly bottom?: Side;
  /** Diagonal stroke (governed together with diagonalUp / diagonalDown). */
  readonly diagonal?: Side;
  /** Vertical stroke between cells of a merged range. */
  readonly vertical?: Side;
  /** Horizontal stroke between cells of a merged range. */
  readonly horizontal?: Side;
  readonly diagonalUp?: boolean;
  readonly diagonalDown?: boolean;
  /** Outline-only flag; defaults to true. */
  readonly outline?: boolean;
}

/** Build an immutable {@link Border}. */
export function makeBorder(opts: Partial<Border> = {}): Border {
  const out: { -readonly [K in keyof Border]: Border[K] } = {};
  if (opts.left !== undefined) out.left = freezeSide(opts.left);
  if (opts.right !== undefined) out.right = freezeSide(opts.right);
  if (opts.top !== undefined) out.top = freezeSide(opts.top);
  if (opts.bottom !== undefined) out.bottom = freezeSide(opts.bottom);
  if (opts.diagonal !== undefined) out.diagonal = freezeSide(opts.diagonal);
  if (opts.vertical !== undefined) out.vertical = freezeSide(opts.vertical);
  if (opts.horizontal !== undefined) out.horizontal = freezeSide(opts.horizontal);
  if (opts.diagonalUp !== undefined) out.diagonalUp = opts.diagonalUp;
  if (opts.diagonalDown !== undefined) out.diagonalDown = opts.diagonalDown;
  if (opts.outline !== undefined) out.outline = opts.outline;
  return Object.freeze(out);
}

const freezeSide = (s: Side): Side => (Object.isFrozen(s) ? s : makeSide(s));

/** Default empty side — convenient sentinel for `no edge stroke`. */
export const EMPTY_SIDE: Side = makeSide();
/** Default empty border — every cell starts here until styled otherwise. */
export const DEFAULT_BORDER: Border = makeBorder();
