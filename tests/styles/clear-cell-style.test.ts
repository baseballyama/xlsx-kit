// Tests for clearCellStyle / clearRangeStyle — Excel "Clear Formatting" parity.

import { describe, expect, it } from 'vitest';
import {
  clearCellStyle,
  clearRangeStyle,
  setBold,
  setCellBackgroundColor,
} from '../../../src/styles/cell-style';
import { addWorksheet, createWorkbook } from '../../../src/workbook/workbook';
import { setCell } from '../../../src/worksheet/worksheet';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('clearCellStyle / clearRangeStyle', () => {
  it('resets a styled cell back to styleId 0', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'styled');
    setBold(wb, cellAt(ws, 1, 1));
    expect(cellAt(ws, 1, 1).styleId).not.toBe(0);
    clearCellStyle(wb, cellAt(ws, 1, 1));
    expect(cellAt(ws, 1, 1).styleId).toBe(0);
  });

  it('clears a fill+bold combo back to default', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'styled');
    setBold(wb, cellAt(ws, 1, 1));
    setCellBackgroundColor(wb, cellAt(ws, 1, 1), 'FF0000');
    clearCellStyle(wb, cellAt(ws, 1, 1));
    expect(cellAt(ws, 1, 1).styleId).toBe(0);
  });

  it('range walk resets every styled cell in the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'c');
    setCell(ws, 2, 2, 'd');
    for (const [row, col] of [[1, 1], [1, 2], [2, 1], [2, 2]] as const) {
      setBold(wb, cellAt(ws, row, col));
      expect(cellAt(ws, row, col).styleId).not.toBe(0);
    }
    clearRangeStyle(wb, ws, 'A1:B2');
    for (const [row, col] of [[1, 1], [1, 2], [2, 1], [2, 2]] as const) {
      expect(cellAt(ws, row, col).styleId).toBe(0);
    }
  });

  it('does not materialise cells that do not exist in a sparse range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'only');
    clearRangeStyle(wb, ws, 'A1:Z99');
    // Only A1 should still exist; the rest must not be created.
    expect(ws.rows.get(1)?.size).toBe(1);
    expect(ws.rows.size).toBe(1);
  });

  it('leaves cells outside the range untouched', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'in');
    setCell(ws, 5, 5, 'out');
    setBold(wb, cellAt(ws, 1, 1));
    setBold(wb, cellAt(ws, 5, 5));
    const outBefore = cellAt(ws, 5, 5).styleId;
    clearRangeStyle(wb, ws, 'A1');
    expect(cellAt(ws, 1, 1).styleId).toBe(0);
    expect(cellAt(ws, 5, 5).styleId).toBe(outBefore);
  });
});
