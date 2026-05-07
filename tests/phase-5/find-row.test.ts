// Tests for findRow — header-driven row Array.find.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { findRow, setCell } from '../../src/worksheet/worksheet';

describe('findRow', () => {
  it('returns the first row that satisfies the predicate', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Bob'); setCell(ws, 3, 2, 25);
    setCell(ws, 4, 1, 'Carol'); setCell(ws, 4, 2, 17);
    expect(findRow(ws, 'A1:B4', (r) => r['name'] === 'Bob')).toEqual({ name: 'Bob', age: 25 });
  });

  it('stops at the first match (does not continue iterating)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k'); setCell(ws, 1, 2, 'order');
    setCell(ws, 2, 1, 'x'); setCell(ws, 2, 2, 1);
    setCell(ws, 3, 1, 'x'); setCell(ws, 3, 2, 2);
    setCell(ws, 4, 1, 'x'); setCell(ws, 4, 2, 3);
    expect(findRow(ws, 'A1:B4', (r) => r['k'] === 'x')).toEqual({ k: 'x', order: 1 });
  });

  it('returns undefined when no row matches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    expect(findRow(ws, 'A1:A3', () => false)).toBeUndefined();
  });

  it('returns undefined for an empty data area (header only)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    expect(findRow(ws, 'A1:A1', () => true)).toBeUndefined();
  });

  it('passes the 0-based row index to the predicate', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    setCell(ws, 4, 1, 'c');
    expect(findRow(ws, 'A1:A4', (_row, i) => i === 2)).toEqual({ k: 'c' });
  });
});
