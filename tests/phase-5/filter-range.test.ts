// Tests for filterRange — header-driven row filter in place.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { filterRange, readRangeAsObjects, setCell } from '../../src/worksheet/worksheet';

describe('filterRange', () => {
  it('keeps only rows the predicate returns true for', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Bob'); setCell(ws, 3, 2, 25);
    setCell(ws, 4, 1, 'Carol'); setCell(ws, 4, 2, 17);
    const kept = filterRange(ws, 'A1:B4', (r) => typeof r['age'] === 'number' && r['age'] >= 18);
    expect(kept).toBe(2);
    expect(readRangeAsObjects(ws, 'A1:B4')).toEqual([
      { name: 'Alice', age: 30 },
      { name: 'Bob', age: 25 },
      { name: null, age: null },
    ]);
  });

  it('returns 0 + clears the data area when the predicate rejects every row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    expect(filterRange(ws, 'A1:A3', () => false)).toBe(0);
    expect(readRangeAsObjects(ws, 'A1:A3')).toEqual([{ k: null }, { k: null }]);
  });

  it('returns the original count + leaves data unchanged when predicate keeps every row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    expect(filterRange(ws, 'A1:A3', () => true)).toBe(2);
    expect(readRangeAsObjects(ws, 'A1:A3')).toEqual([{ k: 'a' }, { k: 'b' }]);
  });

  it('preserves multi-column row identity (not just the predicate column)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'role'); setCell(ws, 1, 3, 'active');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 'admin'); setCell(ws, 2, 3, true);
    setCell(ws, 3, 1, 'Bob'); setCell(ws, 3, 2, 'user'); setCell(ws, 3, 3, false);
    setCell(ws, 4, 1, 'Carol'); setCell(ws, 4, 2, 'admin'); setCell(ws, 4, 3, true);
    filterRange(ws, 'A1:C4', (r) => r['active'] === true);
    expect(readRangeAsObjects(ws, 'A1:C4').slice(0, 2)).toEqual([
      { name: 'Alice', role: 'admin', active: true },
      { name: 'Carol', role: 'admin', active: true },
    ]);
  });
});
