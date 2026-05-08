// Tests for removeAllMergedRanges.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import {
  getMergedCells,
  mergeCells,
  removeAllMergedRanges,
  setCell,
} from '../../src/xlsx/worksheet/worksheet';

describe('removeAllMergedRanges', () => {
  it('drops every merge and returns the count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    mergeCells(ws, 'A1:B2');
    mergeCells(ws, 'D5:E6');
    expect(removeAllMergedRanges(ws)).toBe(2);
    expect(getMergedCells(ws)).toEqual([]);
  });

  it('returns 0 when no merges exist', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(removeAllMergedRanges(ws)).toBe(0);
  });

  it('preserves cell values inside the (formerly) merged ranges', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'tl');
    mergeCells(ws, 'A1:B2');
    removeAllMergedRanges(ws);
    expect(ws.rows.get(1)?.get(1)?.value).toBe('tl');
  });
});
