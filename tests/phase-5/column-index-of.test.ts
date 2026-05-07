// Tests for columnIndexOf — header column index (0-based, range-relative).

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { columnIndexOf, setCell } from '../../src/worksheet/worksheet';

describe('columnIndexOf', () => {
  it('returns the 0-based in-range index of a present header', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 1, 2, 'age');
    setCell(ws, 1, 3, 'role');
    expect(columnIndexOf(ws, 'A1:C1', 'age')).toBe(1);
    expect(columnIndexOf(ws, 'A1:C1', 'role')).toBe(2);
  });

  it('returns -1 when the header is not present', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    expect(columnIndexOf(ws, 'A1:A1', 'missing')).toBe(-1);
  });

  it('returns range-relative index (not the absolute worksheet column)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    // Range starts at column C; "name" is the leftmost in-range header.
    setCell(ws, 1, 3, 'name');
    setCell(ws, 1, 4, 'age');
    expect(columnIndexOf(ws, 'C1:D1', 'name')).toBe(0);
    expect(columnIndexOf(ws, 'C1:D1', 'age')).toBe(1);
  });
});
