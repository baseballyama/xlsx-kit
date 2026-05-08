// Tests for getMergedRangeAt / unmergeCellsAt.

import { describe, expect, it } from 'vitest';
import { rangeToString } from '../../src/xlsx/worksheet/cell-range';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import {
  getMergedCells,
  getMergedRangeAt,
  mergeCells,
  unmergeCellsAt,
} from '../../src/xlsx/worksheet/worksheet';

describe('getMergedRangeAt', () => {
  it('returns the range for any cell inside it (top-left, middle, bottom-right)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    mergeCells(ws, 'B2:D4');
    const tl = getMergedRangeAt(ws, 2, 2);
    const mid = getMergedRangeAt(ws, 3, 3);
    const br = getMergedRangeAt(ws, 4, 4);
    expect(tl && rangeToString(tl)).toBe('B2:D4');
    expect(mid && rangeToString(mid)).toBe('B2:D4');
    expect(br && rangeToString(br)).toBe('B2:D4');
  });

  it('returns undefined for cells outside every merge', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    mergeCells(ws, 'A1:B2');
    expect(getMergedRangeAt(ws, 5, 5)).toBeUndefined();
  });

  it('multi-merge sheet returns the matching one', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    mergeCells(ws, 'A1:B2');
    mergeCells(ws, 'D5:E6');
    const a = getMergedRangeAt(ws, 1, 1);
    const b = getMergedRangeAt(ws, 5, 5);
    expect(a && rangeToString(a)).toBe('A1:B2');
    expect(b && rangeToString(b)).toBe('D5:E6');
  });
});

describe('unmergeCellsAt', () => {
  it('drops the merge containing the cell and returns true', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    mergeCells(ws, 'B2:D4');
    expect(unmergeCellsAt(ws, 3, 3)).toBe(true);
    expect(getMergedCells(ws).length).toBe(0);
  });

  it('returns false when no merge covers the cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    mergeCells(ws, 'A1:B2');
    expect(unmergeCellsAt(ws, 5, 5)).toBe(false);
    expect(getMergedCells(ws).length).toBe(1);
  });

  it('only removes the matching merge — leaves other merges alone', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    mergeCells(ws, 'A1:B2');
    mergeCells(ws, 'D5:E6');
    unmergeCellsAt(ws, 1, 1);
    const remaining = getMergedCells(ws);
    expect(remaining.length).toBe(1);
    if (!remaining[0]) throw new Error('expected one remaining merge');
    expect(rangeToString(remaining[0])).toBe('D5:E6');
  });
});
