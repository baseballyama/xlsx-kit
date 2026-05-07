// Tests for columnAggregates — per-column sum/mean/min/max/count.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { columnAggregates, setCell } from '../../src/worksheet/worksheet';

describe('columnAggregates', () => {
  it('computes sum / mean / min / max / count for a numeric column', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'n');
    setCell(ws, 2, 1, 1);
    setCell(ws, 3, 1, 2);
    setCell(ws, 4, 1, 3);
    setCell(ws, 5, 1, 4);
    expect(columnAggregates(ws, 'A1:A5')['n']).toEqual({
      sum: 10,
      mean: 2.5,
      min: 1,
      max: 4,
      count: 4,
      numericCount: 4,
    });
  });

  it('reports NaN for sum/mean/min/max when the column has no numbers', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 's');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    const a = columnAggregates(ws, 'A1:A3')['s'];
    if (!a) throw new Error('s missing');
    expect(a.numericCount).toBe(0);
    expect(a.count).toBe(2);
    expect(a.sum).toBeNaN();
    expect(a.mean).toBeNaN();
    expect(a.min).toBeNaN();
    expect(a.max).toBeNaN();
  });

  it('count covers non-null cells; numericCount only the JS-number ones', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'mixed');
    setCell(ws, 2, 1, 1);
    setCell(ws, 3, 1, 'two');
    setCell(ws, 4, 1, 3);
    // row 5 is empty (null cell)
    setCell(ws, 6, 1, true);
    const m = columnAggregates(ws, 'A1:A6')['mixed'];
    if (!m) throw new Error('mixed missing');
    expect(m.numericCount).toBe(2);
    expect(m.count).toBe(4); // 1, 'two', 3, true (null skipped)
    expect(m.sum).toBe(4);
    expect(m.mean).toBe(2);
  });

  it('produces independent stats for each column in a multi-column range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'x');
    setCell(ws, 1, 2, 'y');
    setCell(ws, 2, 1, 1);
    setCell(ws, 2, 2, 10);
    setCell(ws, 3, 1, 2);
    setCell(ws, 3, 2, 20);
    const r = columnAggregates(ws, 'A1:B3');
    expect(r['x']?.sum).toBe(3);
    expect(r['y']?.sum).toBe(30);
    expect(r['x']?.mean).toBe(1.5);
    expect(r['y']?.mean).toBe(15);
  });
});
