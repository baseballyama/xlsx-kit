// Tests for unhideRow / unhideColumn.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  getColumnDimension,
  getRowDimension,
  hideColumn,
  hideRow,
  setColumnWidth,
  setRowHeight,
  unhideColumn,
  unhideRow,
} from '../../src/worksheet/worksheet';

describe('unhideRow', () => {
  it('drops the hidden flag', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    hideRow(ws, 3);
    expect(getRowDimension(ws, 3)?.hidden).toBe(true);
    unhideRow(ws, 3);
    expect(getRowDimension(ws, 3)?.hidden).toBeUndefined();
  });

  it('preserves height when unhiding', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRowHeight(ws, 3, 30);
    hideRow(ws, 3);
    unhideRow(ws, 3);
    expect(getRowDimension(ws, 3)?.height).toBe(30);
  });

  it('removes the entry when no other fields remain', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    hideRow(ws, 3);
    unhideRow(ws, 3);
    expect(ws.rowDimensions.has(3)).toBe(false);
  });

  it('no-op for rows without a dimension entry', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    unhideRow(ws, 7);
    expect(ws.rowDimensions.size).toBe(0);
  });
});

describe('unhideColumn', () => {
  it('drops the hidden flag', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    hideColumn(ws, 4);
    expect(getColumnDimension(ws, 4)?.hidden).toBe(true);
    unhideColumn(ws, 4);
    expect(getColumnDimension(ws, 4)?.hidden).toBeUndefined();
  });

  it('preserves width when unhiding', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setColumnWidth(ws, 4, 18);
    hideColumn(ws, 4);
    unhideColumn(ws, 4);
    expect(getColumnDimension(ws, 4)?.width).toBe(18);
  });

  it('removes the entry when no other fields remain', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    hideColumn(ws, 4);
    unhideColumn(ws, 4);
    expect(getColumnDimension(ws, 4)).toBeUndefined();
  });
});
