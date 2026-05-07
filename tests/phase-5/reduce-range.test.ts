// Tests for reduceRange — header-driven row reduce.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { reduceRange, setCell } from '../../src/worksheet/worksheet';

describe('reduceRange', () => {
  it('counts data rows when the reducer increments', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    setCell(ws, 4, 1, 'c');
    expect(reduceRange(ws, 'A1:A4', (acc) => acc + 1, 0)).toBe(3);
  });

  it('sums a numeric column by reading the row object', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'n');
    setCell(ws, 2, 1, 1);
    setCell(ws, 3, 1, 2);
    setCell(ws, 4, 1, 3);
    expect(
      reduceRange(ws, 'A1:A4', (acc, row) => acc + (typeof row['n'] === 'number' ? row['n'] : 0), 0),
    ).toBe(6);
  });

  it('finds the max of a numeric column', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'n');
    setCell(ws, 2, 1, 5);
    setCell(ws, 3, 1, 100);
    setCell(ws, 4, 1, 30);
    expect(
      reduceRange(
        ws,
        'A1:A4',
        (acc, row) => (typeof row['n'] === 'number' && row['n'] > acc ? row['n'] : acc),
        Number.NEGATIVE_INFINITY,
      ),
    ).toBe(100);
  });

  it('returns the initial accumulator when the data area is empty', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    expect(reduceRange(ws, 'A1:A1', (acc) => acc + 1, 7)).toBe(7);
  });

  it('passes the row index to the reducer (0-based)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    expect(reduceRange<number[]>(ws, 'A1:A3', (acc, _row, i) => [...acc, i], [])).toEqual([0, 1]);
  });
});
