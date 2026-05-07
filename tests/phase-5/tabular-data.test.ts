// Tests for tabularData — column-oriented header-driven table read.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell, tabularData } from '../../src/worksheet/worksheet';

describe('tabularData', () => {
  it('returns one array per column, in row order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice');
    setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Bob');
    setCell(ws, 3, 2, 25);
    expect(tabularData(ws, 'A1:B3')).toEqual({
      name: ['Alice', 'Bob'],
      age: [30, 25],
    });
  });

  it('returns empty arrays per column when only the header row is in range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    expect(tabularData(ws, 'A1:B1')).toEqual({ a: [], b: [] });
  });

  it('represents missing data cells as null in the per-column array', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    // row 3 col 1 intentionally empty
    setCell(ws, 4, 1, 'c');
    expect(tabularData(ws, 'A1:A4')).toEqual({ k: ['a', null, 'c'] });
  });

  it('coerces non-string header cells to strings (matches readRangeAsObjects)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 42);
    setCell(ws, 1, 2, true);
    setCell(ws, 2, 1, 'a');
    setCell(ws, 2, 2, 'b');
    expect(tabularData(ws, 'A1:B2')).toEqual({ '42': ['a'], true: ['b'] });
  });

  it('on duplicate header names, both columns concat into one array (column store semantics)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'col');
    setCell(ws, 1, 2, 'col');
    setCell(ws, 2, 1, 'left');
    setCell(ws, 2, 2, 'right');
    // Both data columns push into the single 'col' key — distinct from
    // readRangeAsObjects (row store, last-wins).
    expect(tabularData(ws, 'A1:B2')).toEqual({ col: ['left', 'right'] });
  });
});
