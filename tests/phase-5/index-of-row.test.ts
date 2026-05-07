// Tests for indexOfRow — header-driven Array.findIndex for rows.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { indexOfRow, setCell } from '../../src/worksheet/worksheet';

describe('indexOfRow', () => {
  it('returns the 0-based index of the first matching row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 2, 1, 'Alice');
    setCell(ws, 3, 1, 'Bob');
    setCell(ws, 4, 1, 'Carol');
    expect(indexOfRow(ws, 'A1:A4', (r) => r['name'] === 'Bob')).toBe(1);
  });

  it('returns -1 when no row matches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    expect(indexOfRow(ws, 'A1:A3', () => false)).toBe(-1);
  });

  it('returns -1 for an empty data area', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    expect(indexOfRow(ws, 'A1:A1', () => true)).toBe(-1);
  });

  it('short-circuits at the first match (predicate runs once)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    setCell(ws, 4, 1, 'c');
    let calls = 0;
    expect(
      indexOfRow(ws, 'A1:A4', () => {
        calls++;
        return true;
      }),
    ).toBe(0);
    expect(calls).toBe(1);
  });
});
