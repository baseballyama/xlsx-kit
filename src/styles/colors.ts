// Color value object. Mirrors openpyxl/openpyxl/styles/colors.py.
//
// Excel exposes four mutually-exclusive ways to specify a colour:
//   * rgb     — explicit aRGB hex ("AARRGGBB"); 6-hex inputs auto-pad to 00..
//   * indexed — 0-63 index into the legacy COLOR_INDEX palette (plus 64/65
//                reserved for the system fg/bg).
//   * theme   — 0-N index into the workbook's theme colour scheme.
//   * auto    — boolean "use system default".
//
// Per docs/plan/01-architecture.md §5 these are plain readonly objects;
// `makeColor` freezes its result so the Stylesheet pool can dedupe by
// reference equality.

import { OpenXmlSchemaError } from '../utils/exceptions';

/**
 * Colour reference. All fields are optional but Excel expects exactly
 * one of {rgb, indexed, theme, auto} to be set; if none is, the cell
 * inherits the parent style's colour.
 *
 * `tint` modulates the resolved colour; -1 = full black, +1 = full
 * white.
 */
export interface Color {
  /** "AARRGGBB" hex (uppercase). 6-hex inputs are auto-padded with `00` alpha. */
  readonly rgb?: string;
  /** 0..63 → COLOR_INDEX entry. 64 = system foreground, 65 = system background. */
  readonly indexed?: number;
  /** Theme colour index. */
  readonly theme?: number;
  /** "Auto" / system default. */
  readonly auto?: boolean;
  /** Lightness modulation in [-1, 1]. */
  readonly tint?: number;
}

/**
 * Legacy 64-entry palette indexed colours fall back to. Verbatim from
 * openpyxl/openpyxl/styles/colors.py — must not be reordered.
 */
export const COLOR_INDEX: readonly string[] = Object.freeze([
  '00000000',
  '00FFFFFF',
  '00FF0000',
  '0000FF00',
  '000000FF', // 0-4
  '00FFFF00',
  '00FF00FF',
  '0000FFFF',
  '00000000',
  '00FFFFFF', // 5-9
  '00FF0000',
  '0000FF00',
  '000000FF',
  '00FFFF00',
  '00FF00FF', // 10-14
  '0000FFFF',
  '00800000',
  '00008000',
  '00000080',
  '00808000', // 15-19
  '00800080',
  '00008080',
  '00C0C0C0',
  '00808080',
  '009999FF', // 20-24
  '00993366',
  '00FFFFCC',
  '00CCFFFF',
  '00660066',
  '00FF8080', // 25-29
  '000066CC',
  '00CCCCFF',
  '00000080',
  '00FF00FF',
  '00FFFF00', // 30-34
  '0000FFFF',
  '00800080',
  '00800000',
  '00008080',
  '000000FF', // 35-39
  '0000CCFF',
  '00CCFFFF',
  '00CCFFCC',
  '00FFFF99',
  '0099CCFF', // 40-44
  '00FF99CC',
  '00CC99FF',
  '00FFCC99',
  '003366FF',
  '0033CCCC', // 45-49
  '0099CC00',
  '00FFCC00',
  '00FF9900',
  '00FF6600',
  '00666699', // 50-54
  '00969696',
  '00003366',
  '00339966',
  '00003300',
  '00333300', // 55-59
  '00993300',
  '00993366',
  '00333399',
  '00333333', // 60-63
]);

/** Convenience constants — match openpyxl's exports. Inlined rather
 * than indexed off COLOR_INDEX to keep the type system from widening
 * to `string | undefined` on tuple lookup. */
export const BLACK = '00000000';
export const WHITE = '00FFFFFF';
export const BLUE = '000000FF';

const ARGB_RE = /^([A-Fa-f0-9]{8}|[A-Fa-f0-9]{6})$/;

/**
 * Normalise an aRGB hex string. Accepts either 6 or 8 hex digits;
 * 6-digit input is padded to 8 by prefixing `00` (alpha=0 = fully
 * opaque per Excel convention). Returns the canonical uppercase form.
 */
export function normaliseRgb(value: string): string {
  if (typeof value !== 'string' || !ARGB_RE.test(value)) {
    throw new OpenXmlSchemaError(`Color rgb must be 6 or 8 hex digits; got "${value}"`);
  }
  return (value.length === 6 ? `00${value}` : value).toUpperCase();
}

/**
 * Build an immutable {@link Color}. Validates ranges (indexed in [0, 65],
 * tint in [-1, 1]) and normalises rgb hex.
 */
export function makeColor(opts: Partial<Color> = {}): Color {
  const out: Mutable<Color> = {};
  if (opts.rgb !== undefined) out.rgb = normaliseRgb(opts.rgb);
  if (opts.indexed !== undefined) {
    if (!Number.isInteger(opts.indexed) || opts.indexed < 0 || opts.indexed > 65) {
      throw new OpenXmlSchemaError(`Color indexed must be in [0, 65]; got ${opts.indexed}`);
    }
    out.indexed = opts.indexed;
  }
  if (opts.theme !== undefined) {
    if (!Number.isInteger(opts.theme) || opts.theme < 0) {
      throw new OpenXmlSchemaError(`Color theme must be a non-negative integer; got ${opts.theme}`);
    }
    out.theme = opts.theme;
  }
  if (opts.auto !== undefined) out.auto = opts.auto;
  if (opts.tint !== undefined) {
    if (!Number.isFinite(opts.tint) || opts.tint < -1 || opts.tint > 1) {
      throw new OpenXmlSchemaError(`Color tint must be in [-1, 1]; got ${opts.tint}`);
    }
    out.tint = opts.tint;
  }
  return Object.freeze(out);
}

/**
 * Resolve `indexed` references against {@link COLOR_INDEX}. Returns
 * undefined for 64/65 (system fg/bg, not in the palette) or out-of-range.
 */
export function resolveIndexedColor(idx: number): string | undefined {
  return COLOR_INDEX[idx];
}

/** Shortcut for the common opaque solid colour. */
export function rgbColor(hex: string): Color {
  return makeColor({ rgb: hex });
}

/**
 * Compute the relative luminance of an ARGB / RGB hex string per
 * the WCAG 2.x formula. Returns a value in `[0, 1]` where 0 is
 * black and 1 is white. The alpha channel (if present) is ignored.
 */
export function luminance(hex: string): number {
  const rgb = normaliseRgb(hex); // 8-char AARRGGBB upper-case
  const r = Number.parseInt(rgb.slice(2, 4), 16) / 255;
  const g = Number.parseInt(rgb.slice(4, 6), 16) / 255;
  const b = Number.parseInt(rgb.slice(6, 8), 16) / 255;
  const lin = (c: number): number => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
}

/**
 * WCAG contrast ratio between two ARGB hex colors. Returns a value
 * in `[1, 21]`; 1 = identical luminance, 21 = pure black on pure
 * white. The order of arguments doesn't matter.
 */
export function contrastRatio(hexA: string, hexB: string): number {
  const lA = luminance(hexA);
  const lB = luminance(hexB);
  const [hi, lo] = lA >= lB ? [lA, lB] : [lB, lA];
  return (hi + 0.05) / (lo + 0.05);
}

/**
 * Pick the higher-contrast text color (`'FF000000'` black or
 * `'FFFFFFFF'` white) for a background hex. Useful when applying a
 * solid fill and wanting the cell text to stay readable.
 */
export function pickReadableTextColor(backgroundHex: string): 'FF000000' | 'FFFFFFFF' {
  // WCAG midpoint of 0.179 splits "near-black bg → white text"
  // from "lighter bg → black text".
  return luminance(backgroundHex) < 0.179 ? 'FFFFFFFF' : 'FF000000';
}

// Internal mutable mirror used inside `make*` constructors. Never leaks.
type Mutable<T> = { -readonly [P in keyof T]: T[P] };
