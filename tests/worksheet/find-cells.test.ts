// Tests for findCells / findFirstCell / getCellsInRange iteration helpers.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  findCells,
  findFirstCell,
  getCellsInRange,
  setCell,
} from '../../src/worksheet/worksheet';

describe('findCells', () => {
  it('yields populated cells matching the predicate in row-then-column order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 2, 3, 'target');
    setCell(ws, 3, 1, 'target');
    setCell(ws, 1, 5, 'b');
    const matches = [...findCells(ws, (c) => c.value === 'target')];
    expect(matches.map((c) => `${c.row}:${c.col}`)).toEqual(['2:3', '3:1']);
  });

  it('predicate that always returns false yields nothing', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    expect([...findCells(ws, () => false)]).toEqual([]);
  });

  it('empty worksheet → empty iteration', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect([...findCells(ws, () => true)]).toEqual([]);
  });

  it('visits cells with value === null (placeholder cells)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const placeholder = setCell(ws, 1, 1, null);
    placeholder.styleId = 7;
    const matches = [...findCells(ws, (c) => c.styleId === 7)];
    expect(matches.length).toBe(1);
  });
});

describe('findFirstCell', () => {
  it('returns the earliest matching cell in row-then-column order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 5, 2, 'hit');
    setCell(ws, 2, 4, 'hit'); // earlier in iteration order
    const first = findFirstCell(ws, (c) => c.value === 'hit');
    expect(first?.row).toBe(2);
    expect(first?.col).toBe(4);
  });

  it('returns undefined when nothing matches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    expect(findFirstCell(ws, (c) => c.value === 'missing')).toBeUndefined();
  });
});

describe('getCellsInRange', () => {
  it('only yields populated cells inside the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 2, 2, 'b');
    setCell(ws, 5, 5, 'far');
    const inside = [...getCellsInRange(ws, 'A1:C3')];
    expect(inside.map((c) => c.value)).toEqual(['a', 'b']);
  });

  it('does not auto-allocate empty coordinates', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    const drained = [...getCellsInRange(ws, 'A1:E5')];
    expect(drained.length).toBe(1);
    // No new cells beyond A1.
    expect(ws.rows.get(1)?.size).toBe(1);
    expect(ws.rows.get(2)).toBeUndefined();
  });

  it('single-cell range yields exactly that cell when populated', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 3, 4, 'D3');
    expect([...getCellsInRange(ws, 'D3')].map((c) => c.value)).toEqual(['D3']);
    expect([...getCellsInRange(ws, 'D4')]).toEqual([]);
  });
});
