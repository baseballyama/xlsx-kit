// Tests for clearRange / clearAllCells.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import {
  clearAllCells,
  clearRange,
  countCells,
  mergeCells,
  setCell,
  setColumnWidth,
} from '../../src/xlsx/worksheet/worksheet';

describe('clearRange', () => {
  it('removes only cells inside the range and returns the count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 5, 5, 'far');
    expect(clearRange(ws, 'A1:C3')).toBe(2);
    expect(countCells(ws)).toBe(1);
    expect(ws.rows.get(5)?.get(5)?.value).toBe('far');
  });

  it('returns 0 for an empty range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    expect(clearRange(ws, 'C3:E5')).toBe(0);
  });

  it('prunes row maps that go empty', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 3, 1, 'lonely');
    clearRange(ws, 'A3:Z3');
    expect(ws.rows.get(3)).toBeUndefined();
  });

  it('does not touch dimensions / merges / comments', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setColumnWidth(ws, 2, 30);
    mergeCells(ws, 'D5:E6');
    clearRange(ws, 'A1:Z10');
    expect(ws.columnDimensions.get(2)?.width).toBe(30);
    expect(ws.mergedCells.length).toBe(1);
  });
});

describe('clearAllCells', () => {
  it('removes every populated cell and returns the count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 5, 5, 'far');
    expect(clearAllCells(ws)).toBe(3);
    expect(countCells(ws)).toBe(0);
  });

  it('preserves column dimensions / merges / comments', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setColumnWidth(ws, 2, 30);
    mergeCells(ws, 'D5:E6');
    clearAllCells(ws);
    expect(ws.columnDimensions.get(2)?.width).toBe(30);
    expect(ws.mergedCells.length).toBe(1);
  });

  it('resets _appendRowCursor so the next appendRow starts at 1', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 5, 1, 'tail');
    ws._appendRowCursor = 5;
    clearAllCells(ws);
    expect(ws._appendRowCursor).toBe(0);
  });
});
