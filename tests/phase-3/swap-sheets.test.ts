// Tests for swapSheets.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  setActiveSheet,
  sheetNames,
  swapSheets,
} from '../../src/xlsx/workbook/workbook';

describe('swapSheets', () => {
  it('swaps two adjacent sheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    swapSheets(wb, 'A', 'B');
    expect(sheetNames(wb)).toEqual(['B', 'A', 'C']);
  });

  it('swaps two non-adjacent sheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    swapSheets(wb, 'A', 'C');
    expect(sheetNames(wb)).toEqual(['C', 'B', 'A']);
  });

  it('activeSheetIndex follows the swap', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    setActiveSheet(wb, 'A');
    swapSheets(wb, 'A', 'C');
    expect(wb.activeSheetIndex).toBe(2); // 'A' is now at index 2
  });

  it('swap with itself is a no-op', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    swapSheets(wb, 'A', 'A');
    expect(sheetNames(wb)).toEqual(['A', 'B']);
  });

  it('throws when either title is missing', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(() => swapSheets(wb, 'A', 'Missing')).toThrow(/no sheet named "Missing"/);
    expect(() => swapSheets(wb, 'Missing', 'A')).toThrow(/no sheet named "Missing"/);
  });
});
