// Tests for copyCellStyle / cloneCellStyle.

import { describe, expect, it } from 'vitest';
import {
  cloneCellStyle,
  copyCellStyle,
  getCellFill,
  getCellFont,
  getCellNumberFormat,
  setCellAsCurrency,
  setCellFont,
} from '../../../src/xlsx/styles/cell-style';
import { makeColor } from '../../../src/xlsx/styles/colors';
import { makeFill, makePatternFill } from '../../../src/xlsx/styles/fills';
import { makeFont } from '../../../src/xlsx/styles/fonts';
import { setCellFill } from '../../../src/xlsx/styles/cell-style';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/workbook';
import { setCell } from '../../../src/xlsx/worksheet/worksheet';

describe('copyCellStyle (same workbook)', () => {
  it('shares the source styleId without allocating a new xf', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const src = setCell(ws, 1, 1, 'src');
    const dst = setCell(ws, 2, 1, 'dst');
    setCellFont(wb, src, makeFont({ bold: true }));
    const sizeBefore = wb.styles.cellXfs.length;
    copyCellStyle(wb, src, dst);
    expect(dst.styleId).toBe(src.styleId);
    expect(wb.styles.cellXfs.length).toBe(sizeBefore);
    expect(getCellFont(wb, dst).bold).toBe(true);
  });
});

describe('cloneCellStyle (cross-workbook)', () => {
  it('deep-copies font + fill + numberFormat into the target workbook', () => {
    const wbA = createWorkbook();
    const wsA = addWorksheet(wbA, 'A');
    const src = setCell(wsA, 1, 1, 1234.5);
    setCellFont(wbA, src, makeFont({ bold: true, color: makeColor({ rgb: 'FFAA0000' }) }));
    setCellFill(wbA, src, makeFill(makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFAAFFAA' }) })));
    setCellAsCurrency(wbA, src, { symbol: '€', decimals: 2 });

    const wbB = createWorkbook();
    const wsB = addWorksheet(wbB, 'A');
    const dst = setCell(wsB, 1, 1, 1234.5);
    cloneCellStyle(wbA, src, wbB, dst);

    expect(getCellFont(wbB, dst).bold).toBe(true);
    expect(getCellFont(wbB, dst).color?.rgb).toBe('FFAA0000');
    const fill = getCellFill(wbB, dst);
    expect(fill.kind === 'pattern' ? fill.fgColor?.rgb : undefined).toBe('FFAAFFAA');
    expect(getCellNumberFormat(wbB, dst)).toBe('€#,##0.00');
  });

  it('returns the new styleId in the target workbook', () => {
    const wbA = createWorkbook();
    const wsA = addWorksheet(wbA, 'A');
    const src = setCell(wsA, 1, 1, 'x');
    setCellFont(wbA, src, makeFont({ italic: true }));

    const wbB = createWorkbook();
    const wsB = addWorksheet(wbB, 'A');
    const dst = setCell(wsB, 1, 1, 'x');
    const newId = cloneCellStyle(wbA, src, wbB, dst);
    expect(newId).toBe(dst.styleId);
    expect(newId).toBeGreaterThan(0);
  });

  it('reserves cellXfs[0] in the target workbook', () => {
    const wbA = createWorkbook();
    const wsA = addWorksheet(wbA, 'A');
    const src = setCell(wsA, 1, 1, 'x');
    setCellFont(wbA, src, makeFont({ italic: true }));

    const wbB = createWorkbook();
    const wsB = addWorksheet(wbB, 'A');
    const dst = setCell(wsB, 1, 1, 'x');
    cloneCellStyle(wbA, src, wbB, dst);
    // cellXfs[0] is the implicit default; cellXfs[1] is the cloned italic.
    expect(wbB.styles.cellXfs[0]?.fontId).toBe(0);
  });
});
