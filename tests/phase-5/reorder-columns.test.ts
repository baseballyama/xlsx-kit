// Tests for reorderColumns — header-driven column reorder + optional subset.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  readRangeAsObjects,
  reorderColumns,
  setCell,
} from '../../src/worksheet/worksheet';

describe('reorderColumns', () => {
  it('swaps two columns and returns the same-shape range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 1); setCell(ws, 2, 2, 2);
    setCell(ws, 3, 1, 3); setCell(ws, 3, 2, 4);
    const newRange = reorderColumns(ws, 'A1:B3', ['b', 'a']);
    expect(newRange).toBe('A1:B3');
    expect(readRangeAsObjects(ws, newRange)).toEqual([
      { b: 2, a: 1 },
      { b: 4, a: 3 },
    ]);
  });

  it('drops omitted columns (subset reorder)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'role'); setCell(ws, 1, 3, 'flag');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 'admin'); setCell(ws, 2, 3, true);
    setCell(ws, 3, 1, 'Bob'); setCell(ws, 3, 2, 'user'); setCell(ws, 3, 3, false);
    const newRange = reorderColumns(ws, 'A1:C3', ['flag', 'name']);
    expect(newRange).toBe('A1:B3');
    expect(readRangeAsObjects(ws, newRange)).toEqual([
      { flag: true, name: 'Alice' },
      { flag: false, name: 'Bob' },
    ]);
    // Original rightmost column is cleared.
    expect(ws.rows.get(2)?.get(3)?.value).toBeNull();
  });

  it('throws when newOrder mentions a column not in the headers', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b');
    expect(() => reorderColumns(ws, 'A1:B1', ['a', 'missing'])).toThrow(/missing/);
  });

  it('throws when newOrder is empty', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    expect(() => reorderColumns(ws, 'A1:A1', [])).toThrow(/at least one/);
  });

  it('passing the same order is a no-op (data identical, range unchanged)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 1); setCell(ws, 2, 2, 2);
    const newRange = reorderColumns(ws, 'A1:B2', ['a', 'b']);
    expect(newRange).toBe('A1:B2');
    expect(readRangeAsObjects(ws, newRange)).toEqual([{ a: 1, b: 2 }]);
  });
});
