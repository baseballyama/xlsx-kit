// Tests for getSheetIndex — 0-based tab-strip index lookup.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getSheetIndex,
} from '../../src/xlsx/workbook/workbook';

describe('getSheetIndex', () => {
  it('returns the 0-based tab-strip index', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'first');
    addWorksheet(wb, 'second');
    addWorksheet(wb, 'third');
    expect(getSheetIndex(wb, 'first')).toBe(0);
    expect(getSheetIndex(wb, 'second')).toBe(1);
    expect(getSheetIndex(wb, 'third')).toBe(2);
  });

  it('returns -1 when the title is not present', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'present');
    expect(getSheetIndex(wb, 'missing')).toBe(-1);
  });

  it('finds chartsheets by title too', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'data');
    addChartsheet(wb, 'chart');
    expect(getSheetIndex(wb, 'chart')).toBe(1);
  });

  it('returns -1 for an empty workbook', () => {
    const wb = createWorkbook();
    expect(getSheetIndex(wb, 'anything')).toBe(-1);
  });
});
