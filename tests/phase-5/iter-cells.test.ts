// Tests for iterCells — flat per-worksheet cell iterator.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { iterCells, setCell } from '../../src/worksheet/worksheet';

describe('iterCells', () => {
  it('yields every populated cell in row-major order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'c');
    setCell(ws, 2, 3, 'd');
    expect([...iterCells(ws)].map((c) => c.value)).toEqual(['a', 'b', 'c', 'd']);
  });

  it('yields nothing for an empty worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect([...iterCells(ws)]).toEqual([]);
  });

  it('respects the minRow / maxRow / minCol / maxCol bounds', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 2, 2, 'b');
    setCell(ws, 5, 5, 'c');
    expect([...iterCells(ws, { minRow: 2, maxRow: 4 })].map((c) => c.value)).toEqual(['b']);
    expect([...iterCells(ws, { minCol: 5 })].map((c) => c.value)).toEqual(['c']);
  });

  it('walks rows in numerical order even when inserted out of order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 5, 1, 'r5');
    setCell(ws, 2, 1, 'r2');
    setCell(ws, 1, 1, 'r1');
    expect([...iterCells(ws)].map((c) => c.value)).toEqual(['r1', 'r2', 'r5']);
  });
});
