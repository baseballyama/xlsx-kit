// Tests for getWorksheetAsCsv — whole-sheet CSV shortcut over getRangeAsCsv.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { getWorksheetAsCsv } from '../../src/worksheet/csv';
import { setCell } from '../../src/worksheet/worksheet';

describe('getWorksheetAsCsv', () => {
  it('returns the data extent serialised as CSV', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice');
    setCell(ws, 2, 2, 30);
    expect(getWorksheetAsCsv(ws)).toBe('name,age\nAlice,30');
  });

  it('returns "" for an empty worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getWorksheetAsCsv(ws)).toBe('');
  });

  it('uses the data extent (skips empty rows / cols outside the populated area)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    // Sparse cells at corners → extent is A1:C5 (no padding columns past max)
    setCell(ws, 1, 1, 'tl');
    setCell(ws, 5, 3, 'br');
    const csv = getWorksheetAsCsv(ws);
    // Each line has the same number of fields = maxCol - minCol + 1 = 3
    expect(csv.split('\n').map((l) => l.split(',').length)).toEqual([3, 3, 3, 3, 3]);
    expect(csv.split('\n')[0]).toBe('tl,,');
    expect(csv.split('\n')[4]).toBe(',,br');
  });

  it('forwards opts.delimiter to getRangeAsCsv', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    expect(getWorksheetAsCsv(ws, { delimiter: ';' })).toBe('a;b');
  });
});
