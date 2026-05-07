// Tests for readRangeAsObjects — header-driven table read.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { readRangeAsObjects, setCell } from '../../src/worksheet/worksheet';

describe('readRangeAsObjects', () => {
  it('returns an array of objects keyed by header row values', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice');
    setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Bob');
    setCell(ws, 3, 2, 25);
    expect(readRangeAsObjects(ws, 'A1:B3')).toEqual([
      { name: 'Alice', age: 30 },
      { name: 'Bob', age: 25 },
    ]);
  });

  it('returns [] when the range covers only the header row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    expect(readRangeAsObjects(ws, 'A1:B1')).toEqual([]);
  });

  it('coerces non-string header cells to strings; null headers become ""', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 42);
    setCell(ws, 1, 2, true);
    // intentionally leave A1.col 3 empty so the header[2] becomes ""
    setCell(ws, 2, 1, 'x');
    setCell(ws, 2, 2, 'y');
    setCell(ws, 2, 3, 'z');
    expect(readRangeAsObjects(ws, 'A1:C2')).toEqual([{ '42': 'x', true: 'y', '': 'z' }]);
  });

  it("skipEmptyRows: drops data rows where every cell is null", () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    // row 3 intentionally empty
    setCell(ws, 4, 1, 'b');
    expect(readRangeAsObjects(ws, 'A1:A4')).toEqual([{ k: 'a' }, { k: null }, { k: 'b' }]);
    expect(readRangeAsObjects(ws, 'A1:A4', { skipEmptyRows: true })).toEqual([{ k: 'a' }, { k: 'b' }]);
  });

  it('on duplicate header names, last column wins per JS object semantics', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'col');
    setCell(ws, 1, 2, 'col'); // dup
    setCell(ws, 2, 1, 'left');
    setCell(ws, 2, 2, 'right');
    expect(readRangeAsObjects(ws, 'A1:B2')).toEqual([{ col: 'right' }]);
  });
});
