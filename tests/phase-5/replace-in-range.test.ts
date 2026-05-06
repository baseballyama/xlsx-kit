// Tests for replaceInRange.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { replaceInRange, setCell } from '../../src/worksheet/worksheet';

describe('replaceInRange', () => {
  it('only replaces cells inside the range, returning the count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'foo');
    setCell(ws, 1, 2, 'foo');
    setCell(ws, 5, 5, 'foo');
    expect(replaceInRange(ws, 'A1:B2', 'foo', 'bar')).toBe(2);
    expect(ws.rows.get(1)?.get(1)?.value).toBe('bar');
    expect(ws.rows.get(1)?.get(2)?.value).toBe('bar');
    expect(ws.rows.get(5)?.get(5)?.value).toBe('foo');
  });

  it('predicate variant sees only populated cells in the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 5);
    setCell(ws, 1, 2, 15);
    setCell(ws, 5, 5, 25);
    expect(
      replaceInRange(ws, 'A1:B2', (v) => typeof v === 'number' && v >= 10, 0),
    ).toBe(1);
    expect(ws.rows.get(1)?.get(1)?.value).toBe(5);
    expect(ws.rows.get(1)?.get(2)?.value).toBe(0);
    expect(ws.rows.get(5)?.get(5)?.value).toBe(25);
  });

  it('returns 0 when nothing matches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    expect(replaceInRange(ws, 'A1:Z99', 'missing', 'x')).toBe(0);
    expect(ws.rows.get(1)?.get(1)?.value).toBe('a');
  });

  it('does not auto-allocate empty coordinates inside the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    replaceInRange(ws, 'A1:E5', 'x', 'y');
    expect(ws.rows.get(1)?.size).toBe(1);
    expect(ws.rows.get(2)).toBeUndefined();
  });

  it('numeric / boolean cells are skipped when search is a string', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'foo');
    setCell(ws, 2, 1, 42);
    setCell(ws, 3, 1, true);
    expect(replaceInRange(ws, 'A1:A3', 'foo', 'BAR')).toBe(1);
    expect(ws.rows.get(2)?.get(1)?.value).toBe(42);
    expect(ws.rows.get(3)?.get(1)?.value).toBe(true);
  });
});
