// Tests for isWorkbookEmpty — short-circuiting workbook-level emptiness check.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  isWorkbookEmpty,
} from '../../src/xlsx/workbook/workbook';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

describe('isWorkbookEmpty', () => {
  it('returns true for a brand-new workbook (no sheets)', () => {
    expect(isWorkbookEmpty(createWorkbook())).toBe(true);
  });

  it('returns true for a workbook with only empty worksheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    expect(isWorkbookEmpty(wb)).toBe(true);
  });

  it('returns false as soon as any worksheet has a non-null cell value', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(b, 1, 1, 'x');
    expect(isWorkbookEmpty(wb)).toBe(false);
  });

  it('treats chartsheets as not affecting emptiness (no cells to check)', () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    expect(isWorkbookEmpty(wb)).toBe(true);
  });

  it('returns true when the only populated cells hold null values', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, null);
    setCell(ws, 5, 5, null);
    expect(isWorkbookEmpty(wb)).toBe(true);
  });
});
