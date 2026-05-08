// Tests for collapseRowGroup / expandRowGroup / collapseColumnGroup /
// expandColumnGroup outline collapse helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import {
  collapseColumnGroup,
  collapseRowGroup,
  expandColumnGroup,
  expandRowGroup,
  getColumnDimension,
  getRowDimension,
  groupColumns,
  groupRows,
  setRowHeight,
  type Worksheet,
} from '../../src/xlsx/worksheet/worksheet';

describe('collapseRowGroup / expandRowGroup', () => {
  it('collapse hides + flags collapsed; expand reverses', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    groupRows(ws, 2, 5);
    collapseRowGroup(ws, 2, 5);
    for (let r = 2; r <= 5; r++) {
      const d = getRowDimension(ws, r);
      expect(d?.hidden).toBe(true);
      expect(d?.collapsed).toBe(true);
      expect(d?.outlineLevel).toBe(1);
    }
    expandRowGroup(ws, 2, 5);
    for (let r = 2; r <= 5; r++) {
      const d = getRowDimension(ws, r);
      expect(d?.hidden).toBeUndefined();
      expect(d?.collapsed).toBeUndefined();
      expect(d?.outlineLevel).toBe(1);
    }
  });

  it('preserves height + outlineLevel through collapse/expand', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRowHeight(ws, 2, 30);
    groupRows(ws, 2, 2);
    collapseRowGroup(ws, 2, 2);
    expandRowGroup(ws, 2, 2);
    expect(getRowDimension(ws, 2)?.height).toBe(30);
    expect(getRowDimension(ws, 2)?.outlineLevel).toBe(1);
  });

  it('expand on rows that were never collapsed → no-op for empty entries', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expandRowGroup(ws, 2, 4); // no-op
    expect(ws.rowDimensions.size).toBe(0);
  });

  it('rejects bad ranges', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => collapseRowGroup(ws, 5, 2)).toThrow(/invalid row range/);
    expect(() => expandRowGroup(ws, 0, 1)).toThrow(/invalid row range/);
  });
});

describe('collapseColumnGroup / expandColumnGroup', () => {
  it('collapse hides + flags; expand reverses', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    groupColumns(ws, 3, 5);
    collapseColumnGroup(ws, 3, 5);
    for (let c = 3; c <= 5; c++) {
      const d = getColumnDimension(ws, c);
      expect(d?.hidden).toBe(true);
      expect(d?.collapsed).toBe(true);
      expect(d?.outlineLevel).toBe(1);
    }
    expandColumnGroup(ws, 3, 5);
    for (let c = 3; c <= 5; c++) {
      const d = getColumnDimension(ws, c);
      expect(d?.hidden).toBeUndefined();
      expect(d?.collapsed).toBeUndefined();
      expect(d?.outlineLevel).toBe(1);
    }
  });
});

describe('outline collapse round-trip', () => {
  it('collapsed + hidden + outlineLevel survive saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'O');
    groupRows(ws, 2, 4);
    collapseRowGroup(ws, 2, 4);
    groupColumns(ws, 3, 5);
    collapseColumnGroup(ws, 3, 5);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb2.sheets[0]?.sheet;
    if (!sheet || !('rows' in sheet)) throw new Error('expected worksheet');
    const ws2 = sheet as Worksheet;
    const r = getRowDimension(ws2, 3);
    expect(r?.hidden).toBe(true);
    expect(r?.collapsed).toBe(true);
    const c = getColumnDimension(ws2, 4);
    expect(c?.hidden).toBe(true);
    expect(c?.collapsed).toBe(true);
  });
});
