// Tests for isWorksheetEmpty — short-circuiting populated check.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { isWorksheetEmpty, setCell } from '../../src/xlsx/worksheet/worksheet';

describe('isWorksheetEmpty', () => {
  it('returns true for a freshly created sheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(isWorksheetEmpty(ws)).toBe(true);
  });

  it('returns false as soon as a cell holds a non-null value', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'x');
    expect(isWorksheetEmpty(ws)).toBe(false);
  });

  it('returns true when materialised cells exist but every value is null', () => {
    // Cells can be materialised (e.g. by setRange* helpers) yet hold
    // null values — those don't count as "non-empty".
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, null);
    setCell(ws, 2, 2, null);
    expect(isWorksheetEmpty(ws)).toBe(true);
  });

  it("returns false for the empty string '' (it's not null)", () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, '');
    expect(isWorksheetEmpty(ws)).toBe(false);
  });
});
