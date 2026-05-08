// Tests for getActiveSheetTitle — title of wb.activeSheetIndex.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getActiveSheetTitle,
  setActiveSheet,
} from '../../src/xlsx/workbook/workbook';

describe('getActiveSheetTitle', () => {
  it('returns the first sheet title for a default workbook', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'first');
    addWorksheet(wb, 'second');
    expect(getActiveSheetTitle(wb)).toBe('first');
  });

  it('reflects setActiveSheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'first');
    addWorksheet(wb, 'second');
    setActiveSheet(wb, 'second');
    expect(getActiveSheetTitle(wb)).toBe('second');
  });

  it('returns the chartsheet title when the active slot is a chartsheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'data');
    addChartsheet(wb, 'chart');
    setActiveSheet(wb, 'chart');
    expect(getActiveSheetTitle(wb)).toBe('chart');
  });

  it('returns undefined for an empty workbook', () => {
    expect(getActiveSheetTitle(createWorkbook())).toBeUndefined();
  });
});
