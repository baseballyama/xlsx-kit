// Tests for setCellBackgroundColor / clearCellBackground.

import { describe, expect, it } from 'vitest';
import {
  clearCellBackground,
  getCellFill,
  setCellBackgroundColor,
} from '../../../src/xlsx/styles/cell-style';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/workbook';
import { setCell } from '../../../src/xlsx/worksheet/worksheet';

describe('setCellBackgroundColor', () => {
  it('hex string applies a solid pattern fill', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellBackgroundColor(wb, c, 'FFAAFFAA');
    const fill = getCellFill(wb, c);
    expect(fill.kind).toBe('pattern');
    if (fill.kind === 'pattern') {
      expect(fill.patternType).toBe('solid');
      expect(fill.fgColor?.rgb).toBe('FFAAFFAA');
    }
  });

  it('Color partial passes through theme + tint', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellBackgroundColor(wb, c, { theme: 4, tint: 0.4 });
    const fill = getCellFill(wb, c);
    if (fill.kind === 'pattern') {
      expect(fill.fgColor?.theme).toBe(4);
      expect(fill.fgColor?.tint).toBeCloseTo(0.4);
    }
  });

  it('repeated calls with the same color dedup in the fill pool', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const a = setCell(ws, 1, 1, 'a');
    const b = setCell(ws, 2, 1, 'b');
    setCellBackgroundColor(wb, a, 'FFFF0000');
    const beforeLen = wb.styles.fills.length;
    setCellBackgroundColor(wb, b, 'FFFF0000');
    expect(wb.styles.fills.length).toBe(beforeLen);
    expect(a.styleId).toBe(b.styleId);
  });
});

describe('clearCellBackground', () => {
  it('reverts the fill to DEFAULT_EMPTY_FILL', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellBackgroundColor(wb, c, 'FFAAFFAA');
    clearCellBackground(wb, c);
    const fill = getCellFill(wb, c);
    if (fill.kind === 'pattern') {
      expect(fill.patternType).toBeUndefined();
      expect(fill.fgColor).toBeUndefined();
    }
  });
});
