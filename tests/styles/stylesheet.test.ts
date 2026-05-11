import { describe, expect, it } from 'vitest';
import { DEFAULT_BORDER, makeBorder, makeSide } from '../../src/styles/borders';
import { makeColor } from '../../src/styles/colors';
import { DEFAULT_EMPTY_FILL, DEFAULT_GRAY_FILL, makePatternFill } from '../../src/styles/fills';
import { DEFAULT_FONT, makeFont } from '../../src/styles/fonts';
import { BUILTIN_FORMATS_MAX_SIZE } from '../../src/styles/numbers';
import {
  addBorder,
  addCellStyleXf,
  addCellXf,
  addFill,
  addFont,
  addNumFmt,
  type CellXf,
  defaultCellXf,
  getCustomNumFmts,
  makeStylesheet,
} from '../../src/styles/stylesheet';
import { OpenXmlSchemaError } from '../../src/utils/exceptions';

describe('makeStylesheet — defaults', () => {
  it('starts with the Excel-required entries', () => {
    const ss = makeStylesheet();
    expect(ss.fonts).toEqual([DEFAULT_FONT]);
    expect(ss.fills).toEqual([DEFAULT_EMPTY_FILL, DEFAULT_GRAY_FILL]);
    expect(ss.borders).toEqual([DEFAULT_BORDER]);
    expect(ss.numFmts.size).toBe(0);
    expect(ss.cellXfs).toEqual([]);
    expect(ss.cellStyleXfs).toEqual([]);
  });

  it('default-pool keys are pre-registered for dedup', () => {
    const ss = makeStylesheet();
    expect(addFont(ss, DEFAULT_FONT)).toBe(0);
    expect(addFill(ss, DEFAULT_EMPTY_FILL)).toBe(0);
    expect(addFill(ss, DEFAULT_GRAY_FILL)).toBe(1);
    expect(addBorder(ss, DEFAULT_BORDER)).toBe(0);
  });
});

describe('addFont / addFill / addBorder — dedup', () => {
  it('adding the same Font 1000× yields a single pool entry', () => {
    const ss = makeStylesheet();
    const f = makeFont({ name: 'Arial', size: 12, bold: true });
    let id: number | undefined;
    for (let i = 0; i < 1000; i++) {
      const next = addFont(ss, f);
      if (id === undefined) id = next;
      expect(next).toBe(id);
    }
    expect(ss.fonts.length).toBe(2); // default + Arial
  });

  it('different Fonts allocate distinct ids', () => {
    const ss = makeStylesheet();
    const a = addFont(ss, makeFont({ name: 'Arial' }));
    const b = addFont(ss, makeFont({ name: 'Times' }));
    expect(a).not.toBe(b);
    expect(ss.fonts.length).toBe(3); // default + 2
  });

  it('Fonts with the same fields but different insertion order dedupe via stableStringify', () => {
    const ss = makeStylesheet();
    const a = addFont(ss, makeFont({ name: 'Arial', bold: true, size: 11 }));
    // Build via spread in different order — same object shape, same ID expected.
    const b = addFont(ss, makeFont({ size: 11, name: 'Arial', bold: true }));
    expect(a).toBe(b);
  });

  it('Fills + Borders dedupe symmetrically', () => {
    const ss = makeStylesheet();
    const f = makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FF0000' }) });
    expect(addFill(ss, f)).toBe(addFill(ss, f));
    const b = makeBorder({ left: makeSide({ style: 'thin' }) });
    expect(addBorder(ss, b)).toBe(addBorder(ss, b));
  });
});

describe('addNumFmt', () => {
  it('returns the canonical built-in id for built-in codes', () => {
    const ss = makeStylesheet();
    expect(addNumFmt(ss, 'General')).toBe(0);
    expect(addNumFmt(ss, '0%')).toBe(9);
    expect(addNumFmt(ss, 'mm-dd-yy')).toBe(14);
    expect(addNumFmt(ss, '@')).toBe(49);
    // No custom entries created.
    expect(ss.numFmts.size).toBe(0);
  });

  it('allocates ids ≥ 164 for custom codes', () => {
    const ss = makeStylesheet();
    const a = addNumFmt(ss, 'yyyy-mm-dd');
    const b = addNumFmt(ss, '0.0000');
    expect(a).toBe(BUILTIN_FORMATS_MAX_SIZE);
    expect(b).toBe(BUILTIN_FORMATS_MAX_SIZE + 1);
    // Re-adding the same custom code returns the same id.
    expect(addNumFmt(ss, 'yyyy-mm-dd')).toBe(BUILTIN_FORMATS_MAX_SIZE);
  });

  it('getCustomNumFmts surfaces only the custom entries, sorted by id', () => {
    const ss = makeStylesheet();
    addNumFmt(ss, '0.0000');
    addNumFmt(ss, 'yyyy-mm-dd');
    expect(getCustomNumFmts(ss)).toEqual([
      { id: BUILTIN_FORMATS_MAX_SIZE, code: '0.0000' },
      { id: BUILTIN_FORMATS_MAX_SIZE + 1, code: 'yyyy-mm-dd' },
    ]);
  });
});

describe('addCellXf', () => {
  it('uses defaultCellXf as a sane starting point', () => {
    const ss = makeStylesheet();
    const xf = defaultCellXf();
    expect(addCellXf(ss, xf)).toBe(0);
    expect(addCellXf(ss, xf)).toBe(0); // dedup
    expect(ss.cellXfs.length).toBe(1);
  });

  it('different xfs allocate distinct ids', () => {
    const ss = makeStylesheet();
    const fontId = addFont(ss, makeFont({ bold: true }));
    const a = addCellXf(ss, defaultCellXf());
    const b = addCellXf(ss, { fontId, fillId: 0, borderId: 0, numFmtId: 0, applyFont: true });
    expect(a).not.toBe(b);
    expect(ss.cellXfs.length).toBe(2);
  });

  it('rejects out-of-range font / fill / border ids', () => {
    const ss = makeStylesheet();
    expect(() => addCellXf(ss, { fontId: 99, fillId: 0, borderId: 0, numFmtId: 0 } as CellXf)).toThrowError(
      OpenXmlSchemaError,
    );
    expect(() => addCellXf(ss, { fontId: 0, fillId: 99, borderId: 0, numFmtId: 0 } as CellXf)).toThrowError(
      OpenXmlSchemaError,
    );
    expect(() => addCellXf(ss, { fontId: 0, fillId: 0, borderId: 99, numFmtId: 0 } as CellXf)).toThrowError(
      OpenXmlSchemaError,
    );
  });

  it('cellStyleXfs has its own pool', () => {
    const ss = makeStylesheet();
    const xf = defaultCellXf();
    expect(addCellStyleXf(ss, xf)).toBe(0);
    // The two pools are separate — adding the same xf to cellXfs is a
    // different operation that gets its own index space.
    expect(addCellXf(ss, xf)).toBe(0);
    expect(ss.cellStyleXfs.length).toBe(1);
    expect(ss.cellXfs.length).toBe(1);
  });

  it('xfId references must be in range against cellStyleXfs', () => {
    const ss = makeStylesheet();
    addCellStyleXf(ss, defaultCellXf());
    const ok = addCellXf(ss, { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, xfId: 0 });
    expect(typeof ok).toBe('number');
    expect(() => addCellXf(ss, { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, xfId: 99 } as CellXf)).toThrowError(
      OpenXmlSchemaError,
    );
  });

  it('rejects negative or non-integer numFmtId', () => {
    const ss = makeStylesheet();
    expect(() => addCellXf(ss, { fontId: 0, fillId: 0, borderId: 0, numFmtId: -1 } as CellXf)).toThrowError(
      OpenXmlSchemaError,
    );
    expect(() => addCellXf(ss, { fontId: 0, fillId: 0, borderId: 0, numFmtId: 1.5 } as CellXf)).toThrowError(
      OpenXmlSchemaError,
    );
  });

  it('CellXfs with the same logical fields dedupe across alignment / protection too', () => {
    const ss = makeStylesheet();
    const xf1: CellXf = {
      fontId: 0,
      fillId: 0,
      borderId: 0,
      numFmtId: 0,
      alignment: { horizontal: 'left', wrapText: true },
    };
    const xf2: CellXf = {
      fontId: 0,
      fillId: 0,
      borderId: 0,
      numFmtId: 0,
      alignment: { wrapText: true, horizontal: 'left' },
    };
    const a = addCellXf(ss, xf1);
    const b = addCellXf(ss, xf2);
    expect(a).toBe(b);
  });
});
