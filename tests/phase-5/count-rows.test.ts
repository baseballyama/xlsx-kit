// Tests for countRows — header-driven row count, optionally filtered.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { countRows, setCell } from '../../src/worksheet/worksheet';

describe('countRows', () => {
  it('returns the total data-row count when no predicate is given', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    setCell(ws, 4, 1, 'c');
    expect(countRows(ws, 'A1:A4')).toBe(3);
  });

  it('counts only matching rows when a predicate is given', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'n');
    setCell(ws, 2, 1, 1);
    setCell(ws, 3, 1, 2);
    setCell(ws, 4, 1, 3);
    setCell(ws, 5, 1, 4);
    expect(countRows(ws, 'A1:A5', (r) => typeof r['n'] === 'number' && r['n'] >= 3)).toBe(2);
  });

  it('returns 0 for an empty data area', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    expect(countRows(ws, 'A1:A1')).toBe(0);
    expect(countRows(ws, 'A1:A1', () => true)).toBe(0);
  });

  it('returns 0 when the predicate rejects every row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    expect(countRows(ws, 'A1:A3', () => false)).toBe(0);
  });
});
