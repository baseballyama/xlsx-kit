// Tests for setRangeValuesAtAddress — sheet-qualified A1 → 2D write.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  getRangeValuesAtAddress,
  setRangeValuesAtAddress,
} from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('setRangeValuesAtAddress', () => {
  it('writes a rectangular 2D array starting at the address top-left', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    setRangeValuesAtAddress(wb, 'Data!A1:B2', [
      ['a', 'b'],
      ['c', 'd'],
    ]);
    expect(getRangeValuesAtAddress(wb, 'Data!A1:B2')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('honours quoted sheet titles', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Q1 2024');
    setRangeValuesAtAddress(wb, "'Q1 2024'!A1", [['x']]);
    expect(getRangeValuesAtAddress(wb, "'Q1 2024'!A1")).toEqual([['x']]);
  });

  it('skips null / undefined entries (preserves existing cells)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'preserved');
    setRangeValuesAtAddress(wb, 'Data!A1:B2', [
      [null, 'b'],
      ['c', undefined],
    ]);
    expect(ws.rows.get(1)?.get(1)?.value).toBe('preserved');
    expect(ws.rows.get(1)?.get(2)?.value).toBe('b');
    expect(ws.rows.get(2)?.get(1)?.value).toBe('c');
    expect(ws.rows.get(2)?.has(2)).toBe(false);
  });

  it('round-trips through getRangeValuesAtAddress', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'S');
    const data = [
      ['a', 'b', 'c'],
      [1, 2, 3],
      [true, false, null],
    ];
    setRangeValuesAtAddress(wb, 'S!A1:C3', data);
    expect(getRangeValuesAtAddress(wb, 'S!A1:C3')).toEqual([
      ['a', 'b', 'c'],
      [1, 2, 3],
      [true, false, null],
    ]);
  });

  it('throws when the sheet does not exist', () => {
    const wb = createWorkbook();
    expect(() => setRangeValuesAtAddress(wb, 'Missing!A1', [['x']])).toThrow(/sheet/);
  });
});
