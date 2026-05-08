import { describe, expect, it } from 'vitest';
import { makeAlignment } from '../../../src/xlsx/styles/alignment';
import { DEFAULT_BORDER, makeBorder, makeSide } from '../../../src/xlsx/styles/borders';
import { makeColor } from '../../../src/xlsx/styles/colors';
import { DEFAULT_EMPTY_FILL, makePatternFill } from '../../../src/xlsx/styles/fills';
import { DEFAULT_FONT, makeFont } from '../../../src/xlsx/styles/fonts';
import {
  addNamedStyle,
  BUILTIN_NAMED_STYLES,
  ensureBuiltinStyle,
  type NamedStyle,
} from '../../../src/xlsx/styles/named-styles';
import { makeProtection } from '../../../src/xlsx/styles/protection';
import { makeStylesheet } from '../../../src/xlsx/styles/stylesheet';
import { OpenXmlSchemaError } from '../../../src/utils/exceptions';

describe('addNamedStyle', () => {
  it('registers font / fill / border / numFmt and a cellStyleXf', () => {
    const ss = makeStylesheet();
    const xfId = addNamedStyle(ss, {
      name: 'Custom A',
      font: makeFont({ name: 'Arial', size: 11 }),
      fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FF112233' }) }),
      border: makeBorder({ left: makeSide({ style: 'thin' }) }),
      alignment: makeAlignment({ horizontal: 'center' }),
      protection: makeProtection({ locked: true }),
      numberFormat: '0.00',
    });
    expect(typeof xfId).toBe('number');
    expect(ss.cellStyleXfs.length).toBe(1);
    expect(ss.namedStyles?.length).toBe(1);
    expect(ss.namedStyles?.[0]?.name).toBe('Custom A');
    expect(ss.fonts.length).toBe(2); // default + Arial
    expect(ss.fills.length).toBe(3); // default empty + gray125 + new solid
    expect(ss.borders.length).toBe(2);
  });

  it('is idempotent on the style name', () => {
    const ss = makeStylesheet();
    const a = addNamedStyle(ss, { name: 'X', font: DEFAULT_FONT });
    const b = addNamedStyle(ss, { name: 'X', font: makeFont({ bold: true }) });
    // Re-registering returns the cached xfId rather than allocating a new one.
    expect(b).toBe(a);
    expect(ss.namedStyles?.length).toBe(1);
  });

  it('omits apply* flags on the cellStyleXf — the named-style base entry mirrors openpyxl, leaving apply* to the cellXf bridge', () => {
    const ss = makeStylesheet();
    addNamedStyle(ss, {
      name: 'Bold-only',
      font: makeFont({ bold: true }),
    });
    const xf = ss.cellStyleXfs[0];
    expect(xf?.applyFont).toBeUndefined();
    expect(xf?.applyFill).toBeUndefined();
    expect(xf?.applyBorder).toBeUndefined();
    expect(xf?.applyAlignment).toBeUndefined();
    expect(xf?.applyProtection).toBeUndefined();
  });
});

describe('BUILTIN_NAMED_STYLES catalogue', () => {
  it('exports the canonical "Normal" entry pointing at Calibri 12 minor', () => {
    expect(BUILTIN_NAMED_STYLES['Normal']?.builtinId).toBe(0);
    expect(BUILTIN_NAMED_STYLES['Normal']?.font?.name).toBe('Calibri');
    expect(BUILTIN_NAMED_STYLES['Normal']?.font?.size).toBe(12);
    expect(BUILTIN_NAMED_STYLES['Normal']?.font?.scheme).toBe('minor');
  });

  it('has the semantic Good / Bad / Neutral entries with the well-known fills', () => {
    expect(BUILTIN_NAMED_STYLES['Good']?.fill).toBeDefined();
    expect(BUILTIN_NAMED_STYLES['Bad']?.fill).toBeDefined();
    expect(BUILTIN_NAMED_STYLES['Neutral']?.fill).toBeDefined();
  });

  it('Currency / Percent / Comma carry number-format codes', () => {
    expect(BUILTIN_NAMED_STYLES['Comma']?.numberFormat).toBe('#,##0.00');
    expect(BUILTIN_NAMED_STYLES['Currency']?.numberFormat).toBe('"$"#,##0.00');
    expect(BUILTIN_NAMED_STYLES['Percent']?.numberFormat).toBe('0%');
  });

  it('Hyperlink + Followed Hyperlink point at theme colours 10 / 11', () => {
    expect(BUILTIN_NAMED_STYLES['Hyperlink']?.font?.color?.theme).toBe(10);
    expect(BUILTIN_NAMED_STYLES['Followed Hyperlink']?.font?.color?.theme).toBe(11);
    expect(BUILTIN_NAMED_STYLES['Hyperlink']?.font?.underline).toBe('single');
  });

  it('catalogue is frozen so callers cannot mutate it', () => {
    expect(Object.isFrozen(BUILTIN_NAMED_STYLES)).toBe(true);
  });
});

describe('ensureBuiltinStyle', () => {
  it('registers a known built-in and returns its xfId', () => {
    const ss = makeStylesheet();
    const id = ensureBuiltinStyle(ss, 'Good');
    expect(typeof id).toBe('number');
    expect(ss.namedStyles?.[0]?.builtinId).toBe(BUILTIN_NAMED_STYLES['Good']?.builtinId);
  });

  it('is idempotent on the same name', () => {
    const ss = makeStylesheet();
    const a = ensureBuiltinStyle(ss, 'Good');
    const b = ensureBuiltinStyle(ss, 'Good');
    expect(a).toBe(b);
    expect(ss.namedStyles?.length).toBe(1);
  });

  it('rejects unknown names', () => {
    const ss = makeStylesheet();
    expect(() => ensureBuiltinStyle(ss, 'Imaginary Style')).toThrowError(OpenXmlSchemaError);
  });

  it('shares the default Font / Fill / Border pool slots when the spec omits them', () => {
    const ss = makeStylesheet();
    // A spec that only sets the font (e.g. Linked Cell) should leave the
    // default fill / border pool entries untouched.
    ensureBuiltinStyle(ss, 'Linked Cell');
    const linkedXf = ss.cellStyleXfs[0];
    expect(linkedXf?.fillId).toBe(0); // DEFAULT_EMPTY_FILL
    expect(linkedXf?.borderId).toBe(0); // DEFAULT_BORDER
    expect(ss.fills[0]).toEqual(DEFAULT_EMPTY_FILL);
    expect(ss.borders[0]).toEqual(DEFAULT_BORDER);
  });
});

describe('namedStyle: composing user-defined entries', () => {
  it('user can build their own NamedStyle via the same code path', () => {
    const ss = makeStylesheet();
    const style: NamedStyle = {
      name: 'Highlight',
      font: makeFont({ ...DEFAULT_FONT, bold: true, color: makeColor({ rgb: 'FFFF0000' }) }),
      fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFFFFF00' }) }),
    };
    const xfId = addNamedStyle(ss, style);
    expect(typeof xfId).toBe('number');
    expect(ss.namedStyles?.[0]?.name).toBe('Highlight');
    // builtinId is absent on user-defined styles.
    expect(ss.namedStyles?.[0]?.builtinId).toBeUndefined();
  });
});
