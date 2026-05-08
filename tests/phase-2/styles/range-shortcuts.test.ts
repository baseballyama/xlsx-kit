// Tests for setRangeBackgroundColor / setRangeFont / setRangeNumberFormat shortcuts.

import { describe, expect, it } from 'vitest';
import {
  getCellFill,
  getCellFont,
  getCellNumberFormat,
  setRangeBackgroundColor,
  setRangeFont,
  setRangeNumberFormat,
} from '../../../src/xlsx/styles/cell-style';
import { makeFont } from '../../../src/xlsx/styles/fonts';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/workbook';
import { setCell } from '../../../src/xlsx/worksheet/worksheet';

describe('setRangeBackgroundColor', () => {
  it('every cell in the range gets the same solid fill', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'c');
    setRangeBackgroundColor(wb, ws, 'A1:B2', 'FFAAFFAA');
    for (const r of [1, 2]) {
      for (const c of [1, 2]) {
        const cell = ws.rows.get(r)?.get(c);
        if (!cell) throw new Error(`expected cell at (${r}, ${c})`);
        const fill = getCellFill(wb, cell);
        if (fill.kind !== 'pattern') throw new Error('expected pattern fill');
        expect(fill.fgColor?.rgb).toBe('FFAAFFAA');
      }
    }
  });

  it('Color partial passes through', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setRangeBackgroundColor(wb, ws, 'A1', { theme: 4, tint: 0.4 });
    const cell = ws.rows.get(1)?.get(1);
    if (!cell) throw new Error('expected A1');
    const fill = getCellFill(wb, cell);
    if (fill.kind === 'pattern') {
      expect(fill.fgColor?.theme).toBe(4);
      expect(fill.fgColor?.tint).toBeCloseTo(0.4);
    }
  });
});

describe('setRangeFont', () => {
  it('applies the same font to every cell in the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setRangeFont(wb, ws, 'A1:B1', makeFont({ bold: true, name: 'Arial' }));
    for (const c of [1, 2]) {
      const cell = ws.rows.get(1)?.get(c);
      if (!cell) throw new Error(`expected cell at (1, ${c})`);
      const f = getCellFont(wb, cell);
      expect(f.bold).toBe(true);
      expect(f.name).toBe('Arial');
    }
  });
});

describe('setRangeNumberFormat', () => {
  it('applies the same format-code to every cell in the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1.5);
    setCell(ws, 2, 1, 2.5);
    setRangeNumberFormat(wb, ws, 'A1:A2', '0.00');
    for (const r of [1, 2]) {
      const cell = ws.rows.get(r)?.get(1);
      if (!cell) throw new Error(`expected A${r}`);
      expect(getCellNumberFormat(wb, cell)).toBe('0.00');
    }
  });
});
