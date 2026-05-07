// Tests for getRangeValuesAtAddress — sheet-qualified A1 → 2D values.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  getRangeValuesAtAddress,
} from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('getRangeValuesAtAddress', () => {
  it('reads a rectangular range as a 2D array', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'c');
    setCell(ws, 2, 2, 'd');
    expect(getRangeValuesAtAddress(wb, 'Data!A1:B2')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('returns a 2D array for single-cell addresses too', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'x');
    expect(getRangeValuesAtAddress(wb, 'Data!A1')).toEqual([['x']]);
  });

  it('represents empty cells as null', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'a');
    // B1 / A2 / B2 intentionally empty
    expect(getRangeValuesAtAddress(wb, 'Data!A1:B2')).toEqual([
      ['a', null],
      [null, null],
    ]);
  });

  it('handles quoted sheet titles', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Q1 2024');
    setCell(ws, 1, 1, 'q');
    expect(getRangeValuesAtAddress(wb, "'Q1 2024'!A1")).toEqual([['q']]);
  });

  it('throws when the sheet does not exist', () => {
    const wb = createWorkbook();
    expect(() => getRangeValuesAtAddress(wb, 'Missing!A1:B2')).toThrow(/sheet/);
  });
});
