// Tests for getCellsInRow / getCellsInColumn — enumerate populated cells.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import {
  getCellsInColumn,
  getCellsInRow,
  setCell,
} from '../../src/xlsx/worksheet/worksheet';

describe('getCellsInRow / getCellsInColumn', () => {
  it('row: returns populated cells in column order, skipping gaps', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    // gap at col 2
    setCell(ws, 1, 3, 'c');
    setCell(ws, 1, 5, 'e');
    const cells = getCellsInRow(ws, 1);
    expect(cells.map((c) => c.value)).toEqual(['a', 'c', 'e']);
    expect(cells.map((c) => c.col)).toEqual([1, 3, 5]);
  });

  it('row: returns [] for an empty / absent row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getCellsInRow(ws, 1)).toEqual([]);
    setCell(ws, 5, 1, 'x'); // populate row 5; row 1 still empty
    expect(getCellsInRow(ws, 1)).toEqual([]);
  });

  it('column: returns populated cells in row order, skipping gaps', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 2, 'r1');
    // gap at row 2
    setCell(ws, 3, 2, 'r3');
    setCell(ws, 7, 2, 'r7');
    // adjacent column noise
    setCell(ws, 2, 1, 'noise');
    const cells = getCellsInColumn(ws, 2);
    expect(cells.map((c) => c.value)).toEqual(['r1', 'r3', 'r7']);
    expect(cells.map((c) => c.row)).toEqual([1, 3, 7]);
  });

  it('column: returns [] when no cell occupies the column', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 3, 'c');
    expect(getCellsInColumn(ws, 2)).toEqual([]);
  });

  it('column: walks rows in numerical order even when inserted out of sequence', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 5, 1, 'r5');
    setCell(ws, 2, 1, 'r2');
    setCell(ws, 8, 1, 'r8');
    expect(getCellsInColumn(ws, 1).map((c) => c.value)).toEqual(['r2', 'r5', 'r8']);
  });
});
