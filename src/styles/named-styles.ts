// NamedStyle value object + curated built-in catalogue. Mirrors
// openpyxl/openpyxl/styles/named_styles.py + styles/builtins.py.
//
// A `NamedStyle` is a label that bundles Font + Fill + Border +
// Alignment + Protection + number-format string. Cells reference it
// via the cellStyleXfs pool. This module:
//   1. defines the value type;
//   2. wires `addNamedStyle` (registers sub-objects in the Stylesheet,
//      allocates a cellStyleXf, then attaches the name);
//   3. ships a curated subset of Excel's "Cell Styles" gallery as
//      plain-object specs and an `ensureBuiltinStyle` ergonomics
//      function that registers one on demand.
//
// The Accent1..6 with 20/40/60% variants and `pandas_highlight` are
// deferred — they balloon the bundle without earning their keep on
// the read/write hot path.

import { OpenXmlSchemaError } from '../utils/exceptions';
import type { Alignment } from './alignment';
import type { Border } from './borders';
import { makeColor } from './colors';
import type { Fill } from './fills';
import { makePatternFill } from './fills';
import type { Font } from './fonts';
import { makeFont } from './fonts';
import { addBorder, addCellStyleXf, addFill, addFont, addNumFmt, type CellXf, type Stylesheet } from './stylesheet';

export interface NamedStyle {
  readonly name: string;
  /** OOXML built-in id (0..N); absent for user-defined styles. */
  readonly builtinId?: number;
  readonly customBuiltin?: boolean;
  readonly hidden?: boolean;
  readonly iLevel?: number;
  readonly font?: Font;
  readonly fill?: Fill;
  readonly border?: Border;
  readonly alignment?: Alignment;
  readonly protection?: Protection;
  readonly numberFormat?: string;
}

import type { Protection } from './protection';

/**
 * NamedStyle as it sits inside the Stylesheet (an `<cellStyles>` entry).
 * Captures the resolved cellStyleXfs index so the writer can emit the
 * `<cellStyle xfId="…" name="…" builtinId="…"/>` element directly.
 */
export interface StylesheetNamedStyle {
  readonly name: string;
  readonly xfId: number;
  readonly builtinId?: number;
  readonly customBuiltin?: boolean;
  readonly hidden?: boolean;
  readonly iLevel?: number;
}

/**
 * Register a NamedStyle on the Stylesheet:
 *  1. add Font / Fill / Border / NumberFormat to their pools
 *  2. push a CellXf with apply* flags into cellStyleXfs
 *  3. append a {name, xfId, builtinId} entry into the workbook's
 *     namedStyles list (caller-managed; this function returns the
 *     xfId so callers can connect the dots)
 * Idempotent on (name): re-registering by the same name returns the
 * cached xfId.
 */
export function addNamedStyle(ss: Stylesheet, style: NamedStyle): number {
  const cached = ss._namedStyleByName?.get(style.name);
  if (cached !== undefined) return cached.xfId;

  const fontId = style.font !== undefined ? addFont(ss, style.font) : 0;
  const fillId = style.fill !== undefined ? addFill(ss, style.fill) : 0;
  const borderId = style.border !== undefined ? addBorder(ss, style.border) : 0;
  const numFmtId = style.numberFormat !== undefined ? addNumFmt(ss, style.numberFormat) : 0;

  const xf: CellXf = {
    fontId,
    fillId,
    borderId,
    numFmtId,
    ...(style.alignment !== undefined ? { alignment: style.alignment, applyAlignment: true } : {}),
    ...(style.protection !== undefined ? { protection: style.protection, applyProtection: true } : {}),
    ...(style.font !== undefined ? { applyFont: true } : {}),
    ...(style.fill !== undefined ? { applyFill: true } : {}),
    ...(style.border !== undefined ? { applyBorder: true } : {}),
    ...(style.numberFormat !== undefined ? { applyNumberFormat: true } : {}),
  };
  const xfId = addCellStyleXf(ss, xf);

  if (ss._namedStyleByName === undefined) ss._namedStyleByName = new Map();
  if (ss.namedStyles === undefined) ss.namedStyles = [];
  const entry: StylesheetNamedStyle = {
    name: style.name,
    xfId,
    ...(style.builtinId !== undefined ? { builtinId: style.builtinId } : {}),
    ...(style.customBuiltin !== undefined ? { customBuiltin: style.customBuiltin } : {}),
    ...(style.hidden !== undefined ? { hidden: style.hidden } : {}),
    ...(style.iLevel !== undefined ? { iLevel: style.iLevel } : {}),
  };
  ss.namedStyles.push(entry);
  ss._namedStyleByName.set(style.name, entry);
  return xfId;
}

// ---- built-in catalogue ----------------------------------------------------
//
// Common ground: the gallery's body fonts are all Calibri 12, family 2,
// minor scheme, theme-1 colour. Specific styles override font / fill /
// border / numberFormat as needed.

const BODY_FONT: Font = makeFont({
  name: 'Calibri',
  family: 2,
  size: 12,
  color: makeColor({ theme: 1 }),
  scheme: 'minor',
});

const fillRgb = (rgb: string, fg = true): Fill =>
  makePatternFill(
    fg ? { patternType: 'solid', bgColor: makeColor({ rgb }), fgColor: makeColor({ rgb }) } : { patternType: 'solid' },
  );

/** Curated subset of openpyxl's `styles` dict. Keys match the user-visible names. */
export const BUILTIN_NAMED_STYLES: Readonly<Record<string, NamedStyle>> = Object.freeze({
  // -------- core ----------------------------------------------------------
  Normal: { name: 'Normal', builtinId: 0, font: BODY_FONT },

  // -------- semantic state ------------------------------------------------
  Good: {
    name: 'Good',
    builtinId: 26,
    font: makeFont({ ...BODY_FONT, color: makeColor({ rgb: 'FF006100' }) }),
    fill: fillRgb('FFC6EFCE'),
  },
  Bad: {
    name: 'Bad',
    builtinId: 27,
    font: makeFont({ ...BODY_FONT, color: makeColor({ rgb: 'FF9C0006' }) }),
    fill: fillRgb('FFFFC7CE'),
  },
  Neutral: {
    name: 'Neutral',
    builtinId: 28,
    font: makeFont({ ...BODY_FONT, color: makeColor({ rgb: 'FF9C5700' }) }),
    fill: fillRgb('FFFFEB9C'),
  },

  // -------- data ----------------------------------------------------------
  Calculation: {
    name: 'Calculation',
    builtinId: 22,
    font: makeFont({ ...BODY_FONT, bold: true, color: makeColor({ rgb: 'FFFA7D00' }) }),
    fill: fillRgb('FFF2F2F2'),
  },
  'Check Cell': {
    name: 'Check Cell',
    builtinId: 23,
    font: makeFont({ ...BODY_FONT, bold: true, color: makeColor({ rgb: 'FFFFFFFF' }) }),
    fill: fillRgb('FFA5A5A5'),
  },
  'Linked Cell': {
    name: 'Linked Cell',
    builtinId: 24,
    font: makeFont({ ...BODY_FONT, color: makeColor({ rgb: 'FFFA7D00' }) }),
  },
  Note: {
    name: 'Note',
    builtinId: 10,
    font: BODY_FONT,
    fill: fillRgb('FFFFFFC0'),
  },
  'Warning Text': {
    name: 'Warning Text',
    builtinId: 11,
    font: makeFont({ ...BODY_FONT, color: makeColor({ rgb: 'FFFF0000' }) }),
  },
  Input: {
    name: 'Input',
    builtinId: 20,
    font: makeFont({ ...BODY_FONT, color: makeColor({ rgb: 'FF3F3F76' }) }),
    fill: fillRgb('FFFFCC99'),
  },
  Output: {
    name: 'Output',
    builtinId: 21,
    font: makeFont({ ...BODY_FONT, bold: true, color: makeColor({ rgb: 'FF3F3F3F' }) }),
    fill: fillRgb('FFF2F2F2'),
  },
  'Explanatory Text': {
    name: 'Explanatory Text',
    builtinId: 53,
    font: makeFont({ ...BODY_FONT, italic: true, color: makeColor({ rgb: 'FF7F7F7F' }) }),
  },

  // -------- titles & headings ---------------------------------------------
  Title: {
    name: 'Title',
    builtinId: 15,
    font: makeFont({ name: 'Cambria', family: 2, size: 18, scheme: 'major', color: makeColor({ theme: 3 }) }),
  },
  'Headline 1': {
    name: 'Headline 1',
    builtinId: 16,
    font: makeFont({ ...BODY_FONT, bold: true, size: 15, color: makeColor({ theme: 3 }) }),
  },
  'Headline 2': {
    name: 'Headline 2',
    builtinId: 17,
    font: makeFont({ ...BODY_FONT, bold: true, size: 13, color: makeColor({ theme: 3 }) }),
  },
  'Headline 3': {
    name: 'Headline 3',
    builtinId: 18,
    font: makeFont({ ...BODY_FONT, bold: true, color: makeColor({ theme: 3 }) }),
  },
  'Headline 4': {
    name: 'Headline 4',
    builtinId: 19,
    font: makeFont({ ...BODY_FONT, bold: true, italic: true, color: makeColor({ theme: 3 }) }),
  },
  Total: {
    name: 'Total',
    builtinId: 25,
    font: makeFont({ ...BODY_FONT, bold: true }),
  },

  // -------- numeric -------------------------------------------------------
  Comma: { name: 'Comma', builtinId: 3, font: BODY_FONT, numberFormat: '#,##0.00' },
  'Comma [0]': { name: 'Comma [0]', builtinId: 6, font: BODY_FONT, numberFormat: '#,##0' },
  Currency: { name: 'Currency', builtinId: 4, font: BODY_FONT, numberFormat: '"$"#,##0.00' },
  'Currency [0]': { name: 'Currency [0]', builtinId: 7, font: BODY_FONT, numberFormat: '"$"#,##0' },
  Percent: { name: 'Percent', builtinId: 5, font: BODY_FONT, numberFormat: '0%' },

  // -------- hyperlinks ----------------------------------------------------
  Hyperlink: {
    name: 'Hyperlink',
    builtinId: 8,
    font: makeFont({ ...BODY_FONT, underline: 'single', color: makeColor({ theme: 10 }) }),
  },
  'Followed Hyperlink': {
    name: 'Followed Hyperlink',
    builtinId: 9,
    font: makeFont({ ...BODY_FONT, underline: 'single', color: makeColor({ theme: 11 }) }),
  },
});

/**
 * Register a built-in style with the supplied Stylesheet (idempotent).
 * Returns the cellStyleXfs index. Throws OpenXmlSchemaError when the
 * name is unknown.
 */
export function ensureBuiltinStyle(ss: Stylesheet, name: keyof typeof BUILTIN_NAMED_STYLES | string): number {
  const spec = BUILTIN_NAMED_STYLES[name as keyof typeof BUILTIN_NAMED_STYLES];
  if (spec === undefined) {
    throw new OpenXmlSchemaError(`ensureBuiltinStyle: unknown built-in style "${String(name)}"`);
  }
  return addNamedStyle(ss, spec);
}
