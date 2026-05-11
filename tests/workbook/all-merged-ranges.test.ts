// Tests for getAllMergedRanges.

import { describe, expect, it } from 'vitest';
import { rangeToString } from '../../src/worksheet/cell-range';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getAllMergedRanges,
} from '../../src/workbook/workbook';
import { mergeCells } from '../../src/worksheet/worksheet';

describe('getAllMergedRanges', () => {
  it('aggregates merges across every worksheet in tab-strip order', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    mergeCells(a, 'A1:B2');
    mergeCells(b, 'C1:D1');
    mergeCells(a, 'E5:F6');
    const out = getAllMergedRanges(wb).map(({ sheet, range }) => `${sheet.title}:${rangeToString(range)}`);
    expect(out).toEqual(['A:A1:B2', 'A:E5:F6', 'B:C1:D1']);
  });

  it('skips chartsheets', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    addChartsheet(wb, 'Chart');
    mergeCells(a, 'A1:B2');
    expect(getAllMergedRanges(wb).length).toBe(1);
  });

  it('empty workbook → empty array', () => {
    const wb = createWorkbook();
    expect(getAllMergedRanges(wb)).toEqual([]);
  });
});
