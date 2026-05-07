// Tests for addColumn — append a header-driven column to a range.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { addColumn, readRangeAsObjects, setCell } from '../../src/worksheet/worksheet';

describe('addColumn', () => {
  it('appends a new column with a fixed value and returns the new range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 2, 1, 'Alice');
    setCell(ws, 3, 1, 'Bob');
    const newRange = addColumn(ws, 'A1:A3', 'active', true);
    expect(newRange).toBe('A1:B3');
    expect(readRangeAsObjects(ws, newRange)).toEqual([
      { name: 'Alice', active: true },
      { name: 'Bob', active: true },
    ]);
  });

  it('appends a column with per-row computed values via a function', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'first'); setCell(ws, 1, 2, 'last');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 'Anderson');
    setCell(ws, 3, 1, 'Bob'); setCell(ws, 3, 2, 'Bell');
    addColumn(ws, 'A1:B3', 'full', (row) => `${row['first']} ${row['last']}`);
    expect(readRangeAsObjects(ws, 'A1:C3')).toEqual([
      { first: 'Alice', last: 'Anderson', full: 'Alice Anderson' },
      { first: 'Bob', last: 'Bell', full: 'Bob Bell' },
    ]);
  });

  it('throws when the new column name already exists in the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 1); setCell(ws, 2, 2, 2);
    expect(() => addColumn(ws, 'A1:B2', 'a', 'x')).toThrow(/already exists/);
  });

  it('handles a header-only range (no data rows to fill)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    const newRange = addColumn(ws, 'A1:A1', 'age', 0);
    expect(newRange).toBe('A1:B1');
    expect(ws.rows.get(1)?.get(2)?.value).toBe('age');
  });
});
