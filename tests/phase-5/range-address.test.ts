// Tests for getRangeAddress — sheet-qualified A1 range string.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { getRangeAddress } from '../../src/xlsx/worksheet/worksheet';

describe('getRangeAddress', () => {
  it('returns the sheet-qualified address for a single-cell ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    expect(getRangeAddress(ws, 'A1')).toBe('Data!A1');
  });

  it('returns the sheet-qualified address for a rectangular range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    expect(getRangeAddress(ws, 'A1:B5')).toBe('Data!A1:B5');
  });

  it('quotes the title when it contains spaces / punctuation', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Q1 2024');
    expect(getRangeAddress(ws, 'A1:B5')).toBe("'Q1 2024'!A1:B5");
  });

  it('passes through any range string (row span / column span)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    expect(getRangeAddress(ws, '1:5')).toBe('S!1:5');
    expect(getRangeAddress(ws, 'A:E')).toBe('S!A:E');
  });
});
