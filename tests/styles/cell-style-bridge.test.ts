import { describe, expect, it } from 'vitest';
import { makeCell } from '../../src/cell/cell';
import { makeBorder, makeSide } from '../../src/styles/borders';
import {
  getCellAlignment,
  getCellBorder,
  getCellFill,
  getCellFont,
  getCellNumberFormat,
  getCellProtection,
  setCellAlignment,
  setCellBorder,
  setCellFill,
  setCellFont,
  setCellNumberFormat,
  setCellProtection,
} from '../../src/styles/cell-style';
import { makePatternFill } from '../../src/styles/fills';
import { DEFAULT_FONT, makeFont } from '../../src/styles/fonts';
import { createWorkbook } from '../../src/workbook/workbook';

describe('cell ↔ stylesheet bridge — read defaults', () => {
  it('returns DEFAULT_FONT / DEFAULT_EMPTY_FILL / DEFAULT_BORDER when styleId is the empty default', () => {
    const wb = createWorkbook();
    const c = makeCell(1, 1, 42);
    expect(getCellFont(wb, c)).toEqual(DEFAULT_FONT);
    expect(getCellFill(wb, c).kind).toBe('pattern');
    expect(getCellBorder(wb, c)).toBeDefined();
    expect(getCellAlignment(wb, c)).toEqual({});
    expect(getCellProtection(wb, c).locked).toBe(true);
    expect(getCellNumberFormat(wb, c)).toBe('General');
  });
});

describe('setCellFont', () => {
  it('allocates a CellXf and updates styleId', () => {
    const wb = createWorkbook();
    const c = makeCell(1, 1, 'hi');
    expect(c.styleId).toBe(0);
    setCellFont(wb, c, makeFont({ bold: true }));
    expect(c.styleId).toBeGreaterThanOrEqual(0);
    const xf = wb.styles.cellXfs[c.styleId];
    expect(xf?.applyFont).toBe(true);
    expect(getCellFont(wb, c).bold).toBe(true);
  });

  it('idempotent — setting the same font twice gives the same styleId', () => {
    const wb = createWorkbook();
    const c1 = makeCell(1, 1, 'a');
    const c2 = makeCell(1, 2, 'b');
    setCellFont(wb, c1, makeFont({ italic: true }));
    setCellFont(wb, c2, makeFont({ italic: true }));
    expect(c1.styleId).toBe(c2.styleId);
    // cellXfs[0] is the implicit default reserved by setCellFont's first
    // call; cellXfs[1] is the italic xf shared by both cells.
    expect(wb.styles.cellXfs.length).toBe(2);
  });

  it('different fonts allocate distinct ids', () => {
    const wb = createWorkbook();
    const c = makeCell(1, 1, 'a');
    setCellFont(wb, c, makeFont({ bold: true }));
    const idA = c.styleId;
    setCellFont(wb, c, makeFont({ italic: true }));
    expect(c.styleId).not.toBe(idA);
    expect(getCellFont(wb, c).italic).toBe(true);
    expect(getCellFont(wb, c).bold).toBeUndefined();
  });
});

describe('setCellFill / setCellBorder / setCellAlignment / setCellProtection', () => {
  it('setCellFill sets applyFill and reads back', () => {
    const wb = createWorkbook();
    const c = makeCell(1, 1, 1);
    const fill = makePatternFill({ patternType: 'solid', fgColor: { rgb: 'FFFF0000' } });
    setCellFill(wb, c, fill);
    expect(wb.styles.cellXfs[c.styleId]?.applyFill).toBe(true);
    expect(getCellFill(wb, c)).toEqual(fill);
  });

  it('setCellBorder sets applyBorder and reads back', () => {
    const wb = createWorkbook();
    const c = makeCell(1, 1, 1);
    const border = makeBorder({ left: makeSide({ style: 'thin' }) });
    setCellBorder(wb, c, border);
    expect(wb.styles.cellXfs[c.styleId]?.applyBorder).toBe(true);
    expect(getCellBorder(wb, c)).toEqual(border);
  });

  it('setCellAlignment sets applyAlignment and reads back', () => {
    const wb = createWorkbook();
    const c = makeCell(1, 1, 1);
    setCellAlignment(wb, c, { horizontal: 'center', vertical: 'top' });
    expect(wb.styles.cellXfs[c.styleId]?.applyAlignment).toBe(true);
    expect(getCellAlignment(wb, c)).toEqual({ horizontal: 'center', vertical: 'top' });
  });

  it('setCellProtection sets applyProtection and reads back', () => {
    const wb = createWorkbook();
    const c = makeCell(1, 1, 1);
    setCellProtection(wb, c, { locked: false, hidden: true });
    expect(wb.styles.cellXfs[c.styleId]?.applyProtection).toBe(true);
    expect(getCellProtection(wb, c)).toEqual({ locked: false, hidden: true });
  });
});

describe('setCellNumberFormat', () => {
  it('built-in code resolves to its canonical id, no custom entry added', () => {
    const wb = createWorkbook();
    const c = makeCell(1, 1, 1);
    setCellNumberFormat(wb, c, 'General');
    expect(wb.styles.numFmts.size).toBe(0);
    expect(getCellNumberFormat(wb, c)).toBe('General');
    expect(wb.styles.cellXfs[c.styleId]?.applyNumberFormat).toBe(true);
  });

  it('custom code is registered and read back', () => {
    const wb = createWorkbook();
    const c = makeCell(1, 1, 1);
    setCellNumberFormat(wb, c, '0.0000');
    expect(wb.styles.numFmts.size).toBe(1);
    expect(getCellNumberFormat(wb, c)).toBe('0.0000');
  });
});

describe('multiple set* calls compose', () => {
  it('font + border keep their respective apply flags', () => {
    const wb = createWorkbook();
    const c = makeCell(1, 1, 1);
    setCellFont(wb, c, makeFont({ bold: true }));
    setCellBorder(wb, c, makeBorder({ left: makeSide({ style: 'thin' }) }));
    const xf = wb.styles.cellXfs[c.styleId];
    expect(xf?.applyFont).toBe(true);
    expect(xf?.applyBorder).toBe(true);
    expect(getCellFont(wb, c).bold).toBe(true);
    expect(getCellBorder(wb, c).left?.style).toBe('thin');
  });
});
