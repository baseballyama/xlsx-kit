// Tests for bulk hide/unhide row + column helpers.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import {
  getColumnDimension,
  getRowDimension,
  hideColumns,
  hideRows,
  setColumnWidth,
  unhideColumns,
  unhideRows,
} from '../../src/xlsx/worksheet/worksheet';

describe('hideRows / unhideRows', () => {
  it('hideRows stamps every row in the range with hidden=true', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    hideRows(ws, 2, 5);
    for (let r = 2; r <= 5; r++) {
      expect(getRowDimension(ws, r)?.hidden).toBe(true);
    }
    expect(getRowDimension(ws, 1)).toBeUndefined();
    expect(getRowDimension(ws, 6)).toBeUndefined();
  });

  it('unhideRows reverses the bulk hide', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    hideRows(ws, 2, 4);
    unhideRows(ws, 2, 4);
    for (let r = 2; r <= 4; r++) {
      expect(getRowDimension(ws, r)).toBeUndefined();
    }
  });

  it('rejects invalid ranges', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => hideRows(ws, 5, 2)).toThrow(/invalid row range/);
    expect(() => unhideRows(ws, 0, 1)).toThrow(/invalid row range/);
  });
});

describe('hideColumns / unhideColumns', () => {
  it('hideColumns stamps every column in the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    hideColumns(ws, 3, 5);
    for (let c = 3; c <= 5; c++) {
      expect(getColumnDimension(ws, c)?.hidden).toBe(true);
    }
  });

  it('unhideColumns reverses the bulk hide while preserving width', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setColumnWidth(ws, 3, 16);
    hideColumns(ws, 3, 3);
    unhideColumns(ws, 3, 3);
    expect(getColumnDimension(ws, 3)?.width).toBe(16);
    expect(getColumnDimension(ws, 3)?.hidden).toBeUndefined();
  });

  it('rejects invalid ranges', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => hideColumns(ws, 5, 2)).toThrow(/invalid column range/);
    expect(() => unhideColumns(ws, 0, 1)).toThrow(/invalid column range/);
  });
});
