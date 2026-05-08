// Tests for setRangeWrapText — bulk wrapText toggle on a range.

import { describe, expect, it } from 'vitest';
import {
  alignCellHorizontal,
  getCellAlignment,
  setRangeWrapText,
} from '../../../src/xlsx/styles/cell-style';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/workbook';
import { setCell } from '../../../src/xlsx/worksheet/worksheet';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('setRangeWrapText', () => {
  it('turns wrap on for every cell in the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'c');
    setCell(ws, 2, 2, 'd');
    setRangeWrapText(wb, ws, 'A1:B2');
    for (const [row, col] of [[1, 1], [1, 2], [2, 1], [2, 2]] as const) {
      expect(getCellAlignment(wb, cellAt(ws, row, col)).wrapText).toBe(true);
    }
  });

  it('turns wrap off when on=false (preserving other alignment fields)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'x');
    setRangeWrapText(wb, ws, 'A1');
    expect(getCellAlignment(wb, cellAt(ws, 1, 1)).wrapText).toBe(true);
    setRangeWrapText(wb, ws, 'A1', false);
    expect(getCellAlignment(wb, cellAt(ws, 1, 1)).wrapText).toBe(false);
  });

  it('preserves a pre-existing horizontal alignment', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'x');
    alignCellHorizontal(wb, cellAt(ws, 1, 1), 'right');
    setRangeWrapText(wb, ws, 'A1');
    const a = getCellAlignment(wb, cellAt(ws, 1, 1));
    expect(a.horizontal).toBe('right');
    expect(a.wrapText).toBe(true);
  });

  it('materialises empty cells in the range so the patch is observable', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    // No setCell — range is empty.
    setRangeWrapText(wb, ws, 'A1:B2');
    for (const [row, col] of [[1, 1], [1, 2], [2, 1], [2, 2]] as const) {
      expect(getCellAlignment(wb, cellAt(ws, row, col)).wrapText).toBe(true);
    }
  });
});
