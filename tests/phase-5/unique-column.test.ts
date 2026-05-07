// Tests for uniqueColumn — pluckColumn + Set dedupe.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell, uniqueColumn } from '../../src/worksheet/worksheet';

describe('uniqueColumn', () => {
  it('dedupes values in first-seen order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'tag');
    setCell(ws, 2, 1, 'red');
    setCell(ws, 3, 1, 'blue');
    setCell(ws, 4, 1, 'red');
    setCell(ws, 5, 1, 'green');
    setCell(ws, 6, 1, 'blue');
    expect(uniqueColumn(ws, 'A1:A6', 'tag')).toEqual(['red', 'blue', 'green']);
  });

  it('preserves null as a distinct value', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    // row 3 col 1 left empty (null)
    setCell(ws, 4, 1, 'a');
    setCell(ws, 5, 1, 'b');
    expect(uniqueColumn(ws, 'A1:A5', 'k')).toEqual(['a', null, 'b']);
  });

  it('returns [] when the data area is empty', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    expect(uniqueColumn(ws, 'A1:A1', 'k')).toEqual([]);
  });

  it('throws when the column is not one of the headers', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 2, 1, 'a');
    expect(() => uniqueColumn(ws, 'A1:A2', 'missing')).toThrow(/missing/);
  });
});
