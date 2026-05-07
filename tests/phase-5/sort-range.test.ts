// Tests for sortRange — header-driven row sort in place.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { readRangeAsObjects, setCell, sortRange } from '../../src/worksheet/worksheet';

describe('sortRange', () => {
  it('sorts data rows ascending by string column (header preserved)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 2, 1, 'Charlie');
    setCell(ws, 3, 1, 'Alice');
    setCell(ws, 4, 1, 'Bob');
    sortRange(ws, 'A1:A4', 'name');
    expect(readRangeAsObjects(ws, 'A1:A4')).toEqual([{ name: 'Alice' }, { name: 'Bob' }, { name: 'Charlie' }]);
  });

  it('sorts data rows ascending by numeric column (auto-detected)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'n');
    setCell(ws, 2, 1, 30);
    setCell(ws, 3, 1, 1);
    setCell(ws, 4, 1, 100);
    setCell(ws, 5, 1, 5);
    sortRange(ws, 'A1:A5', 'n');
    expect(readRangeAsObjects(ws, 'A1:A5')).toEqual([{ n: 1 }, { n: 5 }, { n: 30 }, { n: 100 }]);
  });

  it('opts.descending reverses the sort', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'n');
    setCell(ws, 2, 1, 1);
    setCell(ws, 3, 1, 2);
    setCell(ws, 4, 1, 3);
    sortRange(ws, 'A1:A4', 'n', { descending: true });
    expect(readRangeAsObjects(ws, 'A1:A4')).toEqual([{ n: 3 }, { n: 2 }, { n: 1 }]);
  });

  it('null values sort last (regardless of order)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'n');
    setCell(ws, 2, 1, 2);
    // row 3 col 1 left null
    setCell(ws, 4, 1, 1);
    sortRange(ws, 'A1:A4', 'n');
    const rows = readRangeAsObjects(ws, 'A1:A4');
    expect(rows[0]?.['n']).toBe(1);
    expect(rows[1]?.['n']).toBe(2);
    expect(rows[2]?.['n']).toBeNull();
  });

  it('sorts multi-column rows by the chosen key (others tag along)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Charlie'); setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Alice'); setCell(ws, 3, 2, 25);
    setCell(ws, 4, 1, 'Bob'); setCell(ws, 4, 2, 40);
    sortRange(ws, 'A1:B4', 'name');
    expect(readRangeAsObjects(ws, 'A1:B4')).toEqual([
      { name: 'Alice', age: 25 },
      { name: 'Bob', age: 40 },
      { name: 'Charlie', age: 30 },
    ]);
  });

  it('throws when byColumn is not one of the headers', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 2, 1, 'a');
    expect(() => sortRange(ws, 'A1:A2', 'missing')).toThrow(/missing/);
  });
});
