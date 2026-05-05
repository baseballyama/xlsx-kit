// Tests for freezeRows / freezeColumns / freezePanes / unfreezePanes
// and addAutoFilter / addAutoFilterColumn helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  addAutoFilter,
  addAutoFilterColumn,
  removeAutoFilter,
} from '../../src/worksheet/auto-filter';
import {
  freezeColumns,
  freezePanes,
  freezeRows,
  getFreezePanes,
  setCell,
  unfreezePanes,
  type Worksheet,
} from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('freezeRows / freezeColumns / freezePanes', () => {
  it('freezeRows(1) freezes the first row, ref="A2"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    freezeRows(ws, 1);
    expect(getFreezePanes(ws)).toBe('A2');
  });

  it('freezeColumns(2) freezes A and B, ref="C1"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    freezeColumns(ws, 2);
    expect(getFreezePanes(ws)).toBe('C1');
  });

  it('freezePanes(2, 3) freezes top 2 rows and left 3 cols, ref="D3"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    freezePanes(ws, 2, 3);
    expect(getFreezePanes(ws)).toBe('D3');
  });

  it('unfreezePanes drops the freeze', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    freezeRows(ws, 1);
    expect(getFreezePanes(ws)).toBe('A2');
    unfreezePanes(ws);
    expect(getFreezePanes(ws)).toBeUndefined();
  });

  it('rejects non-positive counts', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => freezeRows(ws, 0)).toThrow();
    expect(() => freezeColumns(ws, -1)).toThrow();
    expect(() => freezePanes(ws, 0, 1)).toThrow();
  });

  it('full save → load round-trip preserves the freeze ref', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    freezePanes(ws, 1, 1);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(getFreezePanes(ws2)).toBe('B2');
  });
});

describe('addAutoFilter / addAutoFilterColumn / removeAutoFilter', () => {
  it('addAutoFilter sets the filter ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    const af = addAutoFilter(ws, 'A1:E10');
    expect(af.ref).toBe('A1:E10');
    expect(af.filterColumns).toEqual([]);
    expect(ws.autoFilter?.ref).toBe('A1:E10');
  });

  it('addAutoFilterColumn appends a value-list filter', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    addAutoFilter(ws, 'A1:E10');
    addAutoFilterColumn(ws, 0, ['Open', 'In Progress']);
    addAutoFilterColumn(ws, 2, ['1', '2', '3'], { blank: true });
    expect(ws.autoFilter?.filterColumns.length).toBe(2);
    expect(ws.autoFilter?.filterColumns[0]?.colId).toBe(0);
    expect(ws.autoFilter?.filterColumns[0]?.values).toEqual(['Open', 'In Progress']);
    expect(ws.autoFilter?.filterColumns[1]?.blank).toBe(true);
  });

  it('addAutoFilterColumn throws if no autoFilter is set', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    expect(() => addAutoFilterColumn(ws, 0, ['x'])).toThrow();
  });

  it('removeAutoFilter drops the autoFilter', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    addAutoFilter(ws, 'A1:C10');
    removeAutoFilter(ws);
    expect(ws.autoFilter).toBeUndefined();
  });

  it('survives a save → load round-trip', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    setCell(ws, 1, 1, 'Status');
    addAutoFilter(ws, 'A1:A5');
    addAutoFilterColumn(ws, 0, ['Open', 'Closed']);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.autoFilter?.ref).toBe('A1:A5');
    expect(ws2.autoFilter?.filterColumns[0]?.values).toEqual(['Open', 'Closed']);
  });
});