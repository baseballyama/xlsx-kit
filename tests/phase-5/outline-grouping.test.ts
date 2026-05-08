// Tests for groupRows / ungroupRows / groupColumns / ungroupColumns.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import {
  getColumnDimension,
  getRowDimension,
  groupColumns,
  groupRows,
  setRowHeight,
  ungroupColumns,
  ungroupRows,
  type Worksheet,
} from '../../src/xlsx/worksheet/worksheet';

describe('groupRows', () => {
  it('stamps outlineLevel onto each row in the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    groupRows(ws, 2, 5);
    for (let r = 2; r <= 5; r++) {
      expect(getRowDimension(ws, r)?.outlineLevel).toBe(1);
    }
    expect(getRowDimension(ws, 1)).toBeUndefined();
    expect(getRowDimension(ws, 6)).toBeUndefined();
  });

  it('respects the optional level parameter (nesting)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    groupRows(ws, 3, 4, 2);
    expect(getRowDimension(ws, 3)?.outlineLevel).toBe(2);
  });

  it('preserves existing height when grouping', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRowHeight(ws, 2, 30);
    groupRows(ws, 2, 2);
    expect(getRowDimension(ws, 2)?.height).toBe(30);
    expect(getRowDimension(ws, 2)?.outlineLevel).toBe(1);
  });

  it('rejects bad level / range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => groupRows(ws, 1, 1, 0)).toThrow(/in \[1, 7\]/);
    expect(() => groupRows(ws, 1, 1, 8)).toThrow(/in \[1, 7\]/);
    expect(() => groupRows(ws, 5, 2)).toThrow(/invalid row range/);
  });
});

describe('ungroupRows', () => {
  it('drops outlineLevel and removes pure-outline RowDimension entries', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    groupRows(ws, 2, 4);
    ungroupRows(ws, 2, 4);
    for (let r = 2; r <= 4; r++) {
      expect(getRowDimension(ws, r)).toBeUndefined();
    }
  });

  it('preserves other RowDimension fields when ungrouping', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRowHeight(ws, 2, 30);
    groupRows(ws, 2, 2);
    ungroupRows(ws, 2, 2);
    expect(getRowDimension(ws, 2)?.height).toBe(30);
    expect(getRowDimension(ws, 2)?.outlineLevel).toBeUndefined();
  });
});

describe('groupColumns / ungroupColumns', () => {
  it('stamps outlineLevel onto each column', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    groupColumns(ws, 3, 5);
    for (let c = 3; c <= 5; c++) {
      expect(getColumnDimension(ws, c)?.outlineLevel).toBe(1);
    }
  });

  it('ungroupColumns removes pure-outline entries entirely', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    groupColumns(ws, 3, 5);
    ungroupColumns(ws, 3, 5);
    for (let c = 3; c <= 5; c++) {
      expect(getColumnDimension(ws, c)).toBeUndefined();
    }
  });

  it('rejects bad level / range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => groupColumns(ws, 1, 1, 0)).toThrow(/in \[1, 7\]/);
    expect(() => groupColumns(ws, 5, 2)).toThrow(/invalid column range/);
  });
});

describe('outline grouping round-trip', () => {
  it('row + column outline levels survive saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'G');
    groupRows(ws, 2, 4);
    groupColumns(ws, 3, 5, 2);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb2.sheets[0]?.sheet;
    if (!sheet || !('rows' in sheet)) throw new Error('expected worksheet');
    const ws2 = sheet as Worksheet;
    expect(getRowDimension(ws2, 3)?.outlineLevel).toBe(1);
    expect(getColumnDimension(ws2, 4)?.outlineLevel).toBe(2);
  });
});
