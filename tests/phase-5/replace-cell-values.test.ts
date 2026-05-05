// Tests for the replaceCellValues find-and-replace helper.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  replaceCellValues,
  setCell,
} from '../../src/worksheet/worksheet';

describe('replaceCellValues — string search', () => {
  it('replaces exact-string matches and returns the count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'foo');
    setCell(ws, 1, 2, 'foo');
    setCell(ws, 2, 1, 'bar');
    expect(replaceCellValues(ws, 'foo', 'baz')).toBe(2);
    expect(ws.rows.get(1)?.get(1)?.value).toBe('baz');
    expect(ws.rows.get(1)?.get(2)?.value).toBe('baz');
    expect(ws.rows.get(2)?.get(1)?.value).toBe('bar');
  });

  it('returns 0 when nothing matches', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'foo');
    expect(replaceCellValues(ws, 'missing', 'x')).toBe(0);
    expect(ws.rows.get(1)?.get(1)?.value).toBe('foo');
  });

  it('skips numeric/boolean cells when search is a string', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'foo');
    setCell(ws, 2, 1, 42);
    setCell(ws, 3, 1, true);
    expect(replaceCellValues(ws, 'foo', 'BAR')).toBe(1);
    expect(ws.rows.get(2)?.get(1)?.value).toBe(42);
    expect(ws.rows.get(3)?.get(1)?.value).toBe(true);
  });

  it('does NOT match substrings — only exact-string equality', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'foo bar');
    expect(replaceCellValues(ws, 'foo', 'X')).toBe(0);
    expect(ws.rows.get(1)?.get(1)?.value).toBe('foo bar');
  });
});

describe('replaceCellValues — predicate search', () => {
  it('replaces every cell whose predicate returns true', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 5);
    setCell(ws, 1, 2, 15);
    setCell(ws, 1, 3, 25);
    const n = replaceCellValues(ws, (v) => typeof v === 'number' && v >= 15, 0);
    expect(n).toBe(2);
    expect(ws.rows.get(1)?.get(1)?.value).toBe(5);
    expect(ws.rows.get(1)?.get(2)?.value).toBe(0);
    expect(ws.rows.get(1)?.get(3)?.value).toBe(0);
  });

  it('passes the Cell as second arg so callers can branch on coordinate', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setCell(ws, 2, 1, 2);
    setCell(ws, 3, 1, 3);
    // Replace only cells in the first row.
    replaceCellValues(ws, (_v, c) => c.row === 1, -1);
    expect(ws.rows.get(1)?.get(1)?.value).toBe(-1);
    expect(ws.rows.get(2)?.get(1)?.value).toBe(2);
  });
});
