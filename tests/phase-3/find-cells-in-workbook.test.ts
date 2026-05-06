// Tests for findCellInWorkbook / findCellsInWorkbook.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  findCellInWorkbook,
  findCellsInWorkbook,
} from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('findCellInWorkbook', () => {
  it('returns the first matching cell across all sheets', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'a1');
    setCell(b, 1, 1, 'target');
    setCell(a, 2, 1, 'target'); // earlier in iter order than B's
    const hit = findCellInWorkbook(wb, (cell) => cell.value === 'target');
    expect(hit?.sheet.title).toBe('A');
    expect(hit?.cell.row).toBe(2);
  });

  it('returns undefined when nothing matches', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(findCellInWorkbook(wb, () => false)).toBeUndefined();
  });

  it('predicate sees the owning sheet as second arg', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'val');
    setCell(b, 1, 1, 'val');
    const hit = findCellInWorkbook(wb, (_c, s) => s.title === 'B');
    expect(hit?.sheet.title).toBe('B');
  });
});

describe('findCellsInWorkbook', () => {
  it('returns every match in iteration order', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'a1');
    setCell(a, 2, 2, 'hit');
    setCell(b, 1, 1, 'hit');
    setCell(b, 2, 2, 'b22');
    const all = findCellsInWorkbook(wb, (c) => c.value === 'hit');
    expect(all.map(({ sheet, cell }) => `${sheet.title}:${cell.row}:${cell.col}`)).toEqual([
      'A:2:2',
      'B:1:1',
    ]);
  });

  it('empty workbook → empty array', () => {
    const wb = createWorkbook();
    expect(findCellsInWorkbook(wb, () => true)).toEqual([]);
  });
});
