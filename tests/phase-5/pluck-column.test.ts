// Tests for pluckColumn — single-column array extraction.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { pluckColumn, setCell } from '../../src/worksheet/worksheet';

describe('pluckColumn', () => {
  it('returns the values of one column in row order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Bob'); setCell(ws, 3, 2, 25);
    expect(pluckColumn(ws, 'A1:B3', 'name')).toEqual(['Alice', 'Bob']);
    expect(pluckColumn(ws, 'A1:B3', 'age')).toEqual([30, 25]);
  });

  it('represents empty cells as null', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    // row 3 col 1 left empty
    setCell(ws, 4, 1, 'c');
    expect(pluckColumn(ws, 'A1:A4', 'k')).toEqual(['a', null, 'c']);
  });

  it('throws when the column is not one of the headers', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 2, 1, 'Alice');
    expect(() => pluckColumn(ws, 'A1:A2', 'missing')).toThrow(/missing/);
  });

  it('returns [] when the range covers only the header row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    expect(pluckColumn(ws, 'A1:A1', 'k')).toEqual([]);
  });
});
