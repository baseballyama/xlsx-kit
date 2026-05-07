// Tests for removeColumn — drop a header-driven column with shift-left.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { readRangeAsObjects, removeColumn, setCell } from '../../src/worksheet/worksheet';

describe('removeColumn', () => {
  it('removes a middle column and shifts the right side left', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'role'); setCell(ws, 1, 3, 'active');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 'admin'); setCell(ws, 2, 3, true);
    setCell(ws, 3, 1, 'Bob'); setCell(ws, 3, 2, 'user'); setCell(ws, 3, 3, false);
    const newRange = removeColumn(ws, 'A1:C3', 'role');
    expect(newRange).toBe('A1:B3');
    expect(readRangeAsObjects(ws, newRange)).toEqual([
      { name: 'Alice', active: true },
      { name: 'Bob', active: false },
    ]);
    // The original right-most column (col 3) is cleared.
    expect(ws.rows.get(1)?.get(3)?.value).toBeNull();
    expect(ws.rows.get(2)?.get(3)?.value).toBeNull();
  });

  it('removes the right-most column without changing other columns', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'flag');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, true);
    const newRange = removeColumn(ws, 'A1:B2', 'flag');
    expect(newRange).toBe('A1:A2');
    expect(readRangeAsObjects(ws, newRange)).toEqual([{ name: 'Alice' }]);
    expect(ws.rows.get(1)?.get(2)?.value).toBeNull();
    expect(ws.rows.get(2)?.get(2)?.value).toBeNull();
  });

  it('throws when the column is not in the header row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'role');
    expect(() => removeColumn(ws, 'A1:B1', 'missing')).toThrow(/missing/);
  });

  it('throws when the range has only one column (would leave zero-column range)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 2, 1, 'Alice');
    expect(() => removeColumn(ws, 'A1:A2', 'name')).toThrow(/only column/);
  });
});
