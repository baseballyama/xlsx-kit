// Tests for getRangeValues / getColumnValues / getRowValues read helpers.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  getColumnValues,
  getRangeValues,
  getRowValues,
  setCell,
  setRangeValues,
} from '../../src/worksheet/worksheet';

describe('getRangeValues', () => {
  it('returns a dense 2-D array with nulls for empty cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 3, 'c');
    setCell(ws, 2, 2, 'b');
    expect(getRangeValues(ws, 'A1:C2')).toEqual([
      ['a', null, 'c'],
      [null, 'b', null],
    ]);
  });

  it('round-trips through setRangeValues', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const data = [
      [1, 2, 3],
      [4, 5, 6],
    ];
    setRangeValues(ws, 'B2:D3', data);
    expect(getRangeValues(ws, 'B2:D3')).toEqual(data);
  });

  it('shape matches the parsed range exactly', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const v = getRangeValues(ws, 'A1:E5');
    expect(v.length).toBe(5);
    expect(v[0]?.length).toBe(5);
    expect(v.every((row) => row.every((c) => c === null))).toBe(true);
  });

  it('handles single-cell ranges', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 3, 4, 42);
    expect(getRangeValues(ws, 'D3')).toEqual([[42]]);
  });
});

describe('getColumnValues', () => {
  it('returns null-padded values across populated rows', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 10);
    setCell(ws, 3, 1, 30);
    setCell(ws, 5, 1, 50);
    expect(getColumnValues(ws, 1)).toEqual([10, null, 30, null, 50]);
  });

  it('respects minRow / maxRow window', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 2, 10);
    setCell(ws, 5, 2, 50);
    expect(getColumnValues(ws, 2, { minRow: 2, maxRow: 4 })).toEqual([null, null, null]);
    expect(getColumnValues(ws, 2, { minRow: 4, maxRow: 5 })).toEqual([null, 50]);
  });

  it('empty worksheet → empty array', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getColumnValues(ws, 1)).toEqual([]);
  });
});

describe('getRowValues', () => {
  it('reads from col 1 through max populated col by default', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 4, 'd');
    expect(getRowValues(ws, 1)).toEqual(['a', null, null, 'd']);
  });

  it('respects minCol / maxCol window', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 2, 4, 'd');
    expect(getRowValues(ws, 2, { minCol: 2, maxCol: 5 })).toEqual([null, null, 'd', null]);
  });

  it('missing row → empty array (no opts)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getRowValues(ws, 7)).toEqual([]);
  });

  it('missing row with explicit window pads with nulls', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getRowValues(ws, 7, { minCol: 1, maxCol: 3 })).toEqual([null, null, null]);
  });
});
