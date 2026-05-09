// Pattern + gradient fill value objects. Mirrors
// openpyxl/openpyxl/styles/fills.py.
//
// In OOXML the cell `<fill>` element wraps exactly one of:
//   * <patternFill patternType="…">[<fgColor>, <bgColor>]</patternFill>
//   * <gradientFill type="linear|path" …>[<stop>, …]</gradientFill>
//
// The TS port models the inner variants as a discriminated union via a `kind`
// tag; the wrapper is reconstructed during XML round-trip (see
// fills.schema.ts). All values are plain readonly + frozen so the Stylesheet
// pool can dedupe.

import { OpenXmlSchemaError } from '../utils/exceptions';
import { type Color, colorToHex, makeColor } from './colors';

// ---- pattern fills ---------------------------------------------------------

/** Predefined Excel pattern types. `'none'` lives here too as an explicit value. */
export type PatternType =
  | 'none'
  | 'solid'
  | 'darkDown'
  | 'darkGray'
  | 'darkGrid'
  | 'darkHorizontal'
  | 'darkTrellis'
  | 'darkUp'
  | 'darkVertical'
  | 'gray0625'
  | 'gray125'
  | 'lightDown'
  | 'lightGray'
  | 'lightGrid'
  | 'lightHorizontal'
  | 'lightTrellis'
  | 'lightUp'
  | 'lightVertical'
  | 'mediumGray';

export const PATTERN_TYPES: ReadonlyArray<PatternType> = Object.freeze([
  'none',
  'solid',
  'darkDown',
  'darkGray',
  'darkGrid',
  'darkHorizontal',
  'darkTrellis',
  'darkUp',
  'darkVertical',
  'gray0625',
  'gray125',
  'lightDown',
  'lightGray',
  'lightGrid',
  'lightHorizontal',
  'lightTrellis',
  'lightUp',
  'lightVertical',
  'mediumGray',
]);

const PATTERN_TYPE_SET: ReadonlySet<string> = new Set(PATTERN_TYPES);

export interface PatternFill {
  readonly kind: 'pattern';
  readonly patternType?: PatternType;
  readonly fgColor?: Color;
  readonly bgColor?: Color;
}

export function makePatternFill(opts: Partial<Omit<PatternFill, 'kind'>> = {}): PatternFill {
  const out: { -readonly [K in keyof PatternFill]: PatternFill[K] } = { kind: 'pattern' };
  if (opts.patternType !== undefined) {
    if (!PATTERN_TYPE_SET.has(opts.patternType)) {
      throw new OpenXmlSchemaError(
        `PatternFill patternType must be one of [${PATTERN_TYPES.join(', ')}]; got "${opts.patternType}"`,
      );
    }
    out.patternType = opts.patternType;
  }
  if (opts.fgColor !== undefined) out.fgColor = freezeColor(opts.fgColor);
  if (opts.bgColor !== undefined) out.bgColor = freezeColor(opts.bgColor);
  return Object.freeze(out);
}

// ---- gradient fills --------------------------------------------------------

export interface GradientStop {
  /** 0..1 ratio along the gradient. */
  readonly position: number;
  readonly color: Color;
}

export type GradientFillType = 'linear' | 'path';

export interface GradientFill {
  readonly kind: 'gradient';
  readonly type: GradientFillType;
  /** Rotation angle (degrees). Ignored when `type === 'path'`. */
  readonly degree?: number;
  /** path-mode insets (0..1 from each edge). */
  readonly left?: number;
  readonly right?: number;
  readonly top?: number;
  readonly bottom?: number;
  /** Colour stops along the gradient. */
  readonly stops: ReadonlyArray<GradientStop>;
}

export function makeGradientStop(position: number, color: Color | Partial<Color>): GradientStop {
  if (!Number.isFinite(position) || position < 0 || position > 1) {
    throw new OpenXmlSchemaError(`GradientStop position must be in [0, 1]; got ${position}`);
  }
  return Object.freeze({ position, color: freezeColor(color) });
}

export function makeGradientFill(opts: Partial<Omit<GradientFill, 'kind'>> = {}): GradientFill {
  const out: { -readonly [K in keyof GradientFill]: GradientFill[K] } = {
    kind: 'gradient',
    type: opts.type ?? 'linear',
    stops: opts.stops?.map((s) => freezeStop(s)) ?? [],
  };
  if (out.type !== 'linear' && out.type !== 'path') {
    throw new OpenXmlSchemaError(`GradientFill type must be "linear" or "path"; got "${String(out.type)}"`);
  }
  if (opts.degree !== undefined) out.degree = opts.degree;
  if (opts.left !== undefined) out.left = opts.left;
  if (opts.right !== undefined) out.right = opts.right;
  if (opts.top !== undefined) out.top = opts.top;
  if (opts.bottom !== undefined) out.bottom = opts.bottom;
  Object.freeze(out.stops);
  return Object.freeze(out);
}

// ---- Fill = PatternFill | GradientFill -------------------------------------

export type Fill = PatternFill | GradientFill;

/**
 * Single-arg constructor that defers to the variant-specific maker based on
 * `kind`. Useful when the caller has a plain object in hand and wants the
 * freeze invariant applied uniformly.
 */
export function makeFill(opts: Partial<PatternFill> | Partial<GradientFill>): Fill {
  if (opts.kind === 'gradient') return makeGradientFill(opts as Partial<Omit<GradientFill, 'kind'>>);
  return makePatternFill(opts as Partial<Omit<PatternFill, 'kind'>>);
}

/** The empty PatternFill — Excel's default cellXf[0] points here. */
export const DEFAULT_EMPTY_FILL: Fill = makePatternFill();
/** The 'gray125' PatternFill — Excel's default cellXf[1]. */
export const DEFAULT_GRAY_FILL: Fill = makePatternFill({ patternType: 'gray125' });

// ---- internal helpers ------------------------------------------------------

const freezeColor = (c: Color | Partial<Color>): Color => (Object.isFrozen(c) ? (c as Color) : makeColor(c));
const freezeStop = (s: GradientStop): GradientStop => (Object.isFrozen(s) ? s : makeGradientStop(s.position, s.color));

const argbToCssHex = (color: Color | undefined): string | undefined => {
  const argb = colorToHex(color);
  return argb ? `#${argb.slice(2)}` : undefined;
};

/**
 * Translate a {@link Fill} to a CSS-property record suitable for HTML preview.
 * `'solid'` PatternFill renders as `background-color`, other pattern types
 * collapse to bgColor (CSS has no built-in equivalent of Excel hatch patterns).
 * GradientFill emits a CSS `background-image` with `linear-gradient(<angle>,
 * …)` for `type='linear'` or `radial-gradient(circle, …)` for `type='path'`.
 *
 * theme/auto colours and unresolvable inputs are skipped (returns `{}`) so
 * callers can spread without overwriting upstream defaults.
 */
export function fillToCss(fill: Fill | undefined): Record<string, string> {
  const css: Record<string, string> = {};
  if (!fill) return css;
  if (fill.kind === 'pattern') {
    if (fill.patternType === 'none' || fill.patternType === undefined) return css;
    if (fill.patternType === 'solid') {
      const fg = argbToCssHex(fill.fgColor);
      if (fg !== undefined) css['background-color'] = fg;
      return css;
    }
    // Non-solid patterns collapse to the bg colour as a coarse approximation;
    // CSS has no native Excel hatch equivalent.
    const bg = argbToCssHex(fill.bgColor) ?? argbToCssHex(fill.fgColor);
    if (bg !== undefined) css['background-color'] = bg;
    return css;
  }
  // Gradient fill — need at least one resolvable stop to emit anything.
  const stopParts: string[] = [];
  for (const s of fill.stops) {
    const hex = argbToCssHex(s.color);
    if (hex === undefined) continue;
    stopParts.push(`${hex} ${(s.position * 100).toFixed(2)}%`);
  }
  if (stopParts.length === 0) return css;
  if (fill.type === 'linear') {
    const angle = fill.degree ?? 0;
    css['background-image'] = `linear-gradient(${angle}deg, ${stopParts.join(', ')})`;
  } else {
    css['background-image'] = `radial-gradient(circle, ${stopParts.join(', ')})`;
  }
  return css;
}
