// Tests for the formatAsHeader range preset.

import { describe, expect, it } from 'vitest';
import {
  formatAsHeader,
  getCellBorder,
  getCellFill,
  getCellFont,
} from '../../../src/styles/cell-style';
import { addWorksheet, createWorkbook } from '../../../src/workbook/workbook';
import { setCell } from '../../../src/worksheet/worksheet';

describe('formatAsHeader', () => {
  it('default: bold white on dark blue + medium bottom border', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'Name');
    setCell(ws, 1, 2, 'Score');
    formatAsHeader(wb, ws, 'A1:B1');
    const c1 = ws.rows.get(1)?.get(1);
    if (!c1) throw new Error('expected A1');
    expect(getCellFont(wb, c1).bold).toBe(true);
    expect(getCellFont(wb, c1).color?.rgb).toBe('FFFFFFFF');
    const fill = getCellFill(wb, c1);
    expect(fill.kind === 'pattern' ? fill.fgColor?.rgb : undefined).toBe('FF305496');
    expect(getCellBorder(wb, c1).bottom?.style).toBe('medium');
  });

  it('respects custom fillColor + fontColor + bold=false', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'Header');
    formatAsHeader(wb, ws, 'A1', {
      fillColor: 'FF222222',
      fontColor: 'FFFFFF00',
      bold: false,
    });
    const c = ws.rows.get(1)?.get(1);
    if (!c) throw new Error('expected A1');
    expect(getCellFont(wb, c).bold).toBeFalsy();
    expect(getCellFont(wb, c).color?.rgb).toBe('FFFFFF00');
    const fill = getCellFill(wb, c);
    expect(fill.kind === 'pattern' ? fill.fgColor?.rgb : undefined).toBe('FF222222');
  });

  it('bottomBorder=false drops the border axis entirely', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'h');
    formatAsHeader(wb, ws, 'A1', { bottomBorder: false });
    const c = ws.rows.get(1)?.get(1);
    if (!c) throw new Error('expected A1');
    expect(getCellBorder(wb, c).bottom?.style).toBeUndefined();
  });

  it('custom bottom-border style + color', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'h');
    formatAsHeader(wb, ws, 'A1', { bottomBorder: 'thick', bottomBorderColor: 'FFFF0000' });
    const c = ws.rows.get(1)?.get(1);
    if (!c) throw new Error('expected A1');
    const b = getCellBorder(wb, c);
    expect(b.bottom?.style).toBe('thick');
    expect(b.bottom?.color?.rgb).toBe('FFFF0000');
  });

  it('auto-allocates blank cells inside the range so they pick up the style', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'A');
    // B1 and C1 are intentionally absent — formatAsHeader should still apply
    // the preset to them because setRangeStyle auto-allocates.
    formatAsHeader(wb, ws, 'A1:C1');
    expect(ws.rows.get(1)?.get(2)).toBeDefined();
    expect(ws.rows.get(1)?.get(3)).toBeDefined();
    const b = ws.rows.get(1)?.get(3);
    if (!b) throw new Error('expected C1');
    expect(getCellFont(wb, b).bold).toBe(true);
  });

  it('repeated calls with identical opts are idempotent on the cellXfs pool', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'h');
    formatAsHeader(wb, ws, 'A1');
    const sizeAfterFirst = wb.styles.cellXfs.length;
    formatAsHeader(wb, ws, 'A1');
    expect(wb.styles.cellXfs.length).toBe(sizeAfterFirst);
  });
});
