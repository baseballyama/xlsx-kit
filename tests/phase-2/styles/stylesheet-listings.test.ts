// Tests for listFonts / listFills / listBorders / listCellXfs / listCellStyleXfs.

import { describe, expect, it } from 'vitest';
import { setCellBorder, setCellFill, setCellFont } from '../../../src/xlsx/styles/cell-style';
import { makeColor } from '../../../src/xlsx/styles/colors';
import { makeBorder, makeSide } from '../../../src/xlsx/styles/borders';
import { makeFill, makePatternFill } from '../../../src/xlsx/styles/fills';
import { makeFont } from '../../../src/xlsx/styles/fonts';
import {
  listBorders,
  listCellStyleXfs,
  listCellXfs,
  listFills,
  listFonts,
} from '../../../src/xlsx/styles/stylesheet';
import { addNamedStyle } from '../../../src/xlsx/styles/named-styles';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/workbook';
import { setCell } from '../../../src/xlsx/worksheet/worksheet';

describe('stylesheet pool listings', () => {
  it('listFonts starts at 1 (DEFAULT_FONT) and grows with adds', () => {
    const wb = createWorkbook();
    expect(listFonts(wb.styles).length).toBe(1);
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellFont(wb, c, makeFont({ bold: true }));
    expect(listFonts(wb.styles).length).toBe(2);
  });

  it('listFills starts at 2 (empty + gray125) and grows with adds', () => {
    const wb = createWorkbook();
    expect(listFills(wb.styles).length).toBe(2);
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellFill(wb, c, makeFill(makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFAA0000' }) })));
    expect(listFills(wb.styles).length).toBe(3);
  });

  it('listBorders starts at 1 (DEFAULT_BORDER) and grows with adds', () => {
    const wb = createWorkbook();
    expect(listBorders(wb.styles).length).toBe(1);
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellBorder(wb, c, makeBorder({ top: makeSide({ style: 'thin' }) }));
    expect(listBorders(wb.styles).length).toBe(2);
  });

  it('listCellXfs reflects add-on-style allocations', () => {
    const wb = createWorkbook();
    expect(listCellXfs(wb.styles).length).toBe(0);
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellFont(wb, c, makeFont({ bold: true }));
    // One reserved default xf + one styled xf.
    expect(listCellXfs(wb.styles).length).toBe(2);
  });

  it('listCellStyleXfs reflects addNamedStyle inserts', () => {
    const wb = createWorkbook();
    expect(listCellStyleXfs(wb.styles).length).toBe(0);
    addNamedStyle(wb.styles, { name: 'My Style', font: makeFont({ italic: true }) });
    expect(listCellStyleXfs(wb.styles).length).toBe(1);
  });
});
