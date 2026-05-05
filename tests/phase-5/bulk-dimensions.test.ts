// Tests for setColumnWidths / setRowHeights bulk helpers.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  getColumnDimension,
  getRowDimension,
  setColumnWidths,
  setRowHeights,
} from '../../src/worksheet/worksheet';

describe('setColumnWidths', () => {
  it('positional array starting at column 1 by default', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setColumnWidths(ws, [12, 16, 20]);
    expect(getColumnDimension(ws, 1)?.width).toBe(12);
    expect(getColumnDimension(ws, 2)?.width).toBe(16);
    expect(getColumnDimension(ws, 3)?.width).toBe(20);
    expect(getColumnDimension(ws, 1)?.customWidth).toBe(true);
  });

  it('positional array with a startCol offset', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setColumnWidths(ws, [10, 11], 5);
    expect(getColumnDimension(ws, 5)?.width).toBe(10);
    expect(getColumnDimension(ws, 6)?.width).toBe(11);
    expect(getColumnDimension(ws, 1)).toBeUndefined();
  });

  it('Record<col, width> assigns sparsely', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setColumnWidths(ws, { 1: 12, 5: 30, 10: 18 });
    expect(getColumnDimension(ws, 1)?.width).toBe(12);
    expect(getColumnDimension(ws, 5)?.width).toBe(30);
    expect(getColumnDimension(ws, 10)?.width).toBe(18);
    expect(getColumnDimension(ws, 2)).toBeUndefined();
  });

  it('skips non-finite / non-integer entries', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setColumnWidths(ws, [12, NaN, 20]);
    expect(getColumnDimension(ws, 1)?.width).toBe(12);
    expect(getColumnDimension(ws, 2)).toBeUndefined();
    expect(getColumnDimension(ws, 3)?.width).toBe(20);
    setColumnWidths(ws, { 0: 5, 1: 14 });
    expect(getColumnDimension(ws, 1)?.width).toBe(14);
  });
});

describe('setRowHeights', () => {
  it('positional array starting at row 1 by default', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRowHeights(ws, [24, 30, 16]);
    expect(getRowDimension(ws, 1)?.height).toBe(24);
    expect(getRowDimension(ws, 2)?.height).toBe(30);
    expect(getRowDimension(ws, 3)?.height).toBe(16);
    expect(getRowDimension(ws, 1)?.customHeight).toBe(true);
  });

  it('Record<row, height> assigns sparsely', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRowHeights(ws, { 5: 22, 7: 30 });
    expect(getRowDimension(ws, 5)?.height).toBe(22);
    expect(getRowDimension(ws, 7)?.height).toBe(30);
    expect(getRowDimension(ws, 6)).toBeUndefined();
  });

  it('startRow offset', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRowHeights(ws, [50, 60], 10);
    expect(getRowDimension(ws, 10)?.height).toBe(50);
    expect(getRowDimension(ws, 11)?.height).toBe(60);
  });
});