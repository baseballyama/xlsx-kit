// Tests for the combined setCellStyle helper.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import {
  getCellAlignment,
  getCellBorder,
  getCellFill,
  getCellFont,
  getCellNumberFormat,
  getCellProtection,
  makeAlignment,
  makeBorder,
  makeColor,
  makeFont,
  makePatternFill,
  makeProtection,
  makeSide,
  setCellStyle,
} from '../../src/styles';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('setCellStyle (combined setter)', () => {
  it('applies font + fill in a single call', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'Hello');
    setCellStyle(wb, c, {
      font: makeFont({ name: 'Arial', size: 14, bold: true }),
      fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFFF00' }) }),
    });
    expect(getCellFont(wb, c)?.name).toBe('Arial');
    expect(getCellFont(wb, c)?.bold).toBe(true);
    expect(getCellFill(wb, c)).toBeDefined();
  });

  it('applies all six axes at once', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 100);
    const thick = makeSide({ style: 'thick', color: makeColor({ rgb: 'FF000000' }) });
    setCellStyle(wb, c, {
      font: makeFont({ size: 12, italic: true }),
      fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFCCCCCC' }) }),
      border: makeBorder({ left: thick, right: thick, top: thick, bottom: thick }),
      alignment: makeAlignment({ horizontal: 'center', vertical: 'center' }),
      protection: makeProtection({ locked: false }),
      numberFormat: '0.00',
    });
    expect(getCellFont(wb, c)?.italic).toBe(true);
    expect(getCellFill(wb, c)).toBeDefined();
    expect(getCellBorder(wb, c)?.left?.style).toBe('thick');
    expect(getCellAlignment(wb, c)?.horizontal).toBe('center');
    expect(getCellProtection(wb, c)?.locked).toBe(false);
    expect(getCellNumberFormat(wb, c)).toBe('0.00');
  });

  it('empty opts is a no-op', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'X');
    const before = c.styleId;
    setCellStyle(wb, c, {});
    expect(c.styleId).toBe(before);
  });

  it('full save → load round-trip preserves all axes', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'Styled');
    setCellStyle(wb, c, {
      font: makeFont({ name: 'Calibri', size: 16, bold: true, color: makeColor({ rgb: 'FFC00000' }) }),
      fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFFFFF00' }) }),
      alignment: makeAlignment({ horizontal: 'center' }),
      numberFormat: '#,##0.00',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const c2 = ws2.rows.get(1)?.get(1);
    if (!c2) throw new Error('cell missing on reload');
    expect(getCellFont(wb2, c2)?.name).toBe('Calibri');
    expect(getCellFont(wb2, c2)?.bold).toBe(true);
    expect(getCellFill(wb2, c2)).toBeDefined();
    expect(getCellAlignment(wb2, c2)?.horizontal).toBe('center');
    expect(getCellNumberFormat(wb2, c2)).toBe('#,##0.00');
  });
});