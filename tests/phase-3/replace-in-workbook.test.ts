// Tests for replaceCellValuesInWorkbook.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  replaceCellValuesInWorkbook,
} from '../../src/xlsx/workbook/workbook';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

describe('replaceCellValuesInWorkbook', () => {
  it('replaces matches across every worksheet (string mode)', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'foo');
    setCell(a, 2, 2, 'foo');
    setCell(b, 1, 1, 'foo');
    setCell(b, 1, 2, 'other');
    expect(replaceCellValuesInWorkbook(wb, 'foo', 'bar')).toBe(3);
    expect(a.rows.get(1)?.get(1)?.value).toBe('bar');
    expect(a.rows.get(2)?.get(2)?.value).toBe('bar');
    expect(b.rows.get(1)?.get(1)?.value).toBe('bar');
    expect(b.rows.get(1)?.get(2)?.value).toBe('other');
  });

  it('predicate variant sees value, cell, and sheet', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'x');
    setCell(b, 1, 1, 'x');
    expect(replaceCellValuesInWorkbook(wb, (_v, _c, s) => s.title === 'B', 'changed')).toBe(1);
    expect(a.rows.get(1)?.get(1)?.value).toBe('x');
    expect(b.rows.get(1)?.get(1)?.value).toBe('changed');
  });

  it('returns 0 when nothing matches', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(replaceCellValuesInWorkbook(wb, 'missing', 'y')).toBe(0);
  });

  it('numeric / boolean cells are skipped in string mode', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 1, 1, 'foo');
    setCell(a, 2, 1, 42);
    setCell(a, 3, 1, true);
    expect(replaceCellValuesInWorkbook(wb, 'foo', 'BAR')).toBe(1);
    expect(a.rows.get(2)?.get(1)?.value).toBe(42);
    expect(a.rows.get(3)?.get(1)?.value).toBe(true);
  });
});
