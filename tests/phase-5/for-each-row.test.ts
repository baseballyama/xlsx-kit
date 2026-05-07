// Tests for forEachRow — header-driven row iteration.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { forEachRow, setCell } from '../../src/worksheet/worksheet';

describe('forEachRow', () => {
  it('invokes the callback once per data row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    setCell(ws, 4, 1, 'c');
    const seen: string[] = [];
    forEachRow(ws, 'A1:A4', (row) => {
      const v = row['k'];
      if (typeof v === 'string') seen.push(v);
    });
    expect(seen).toEqual(['a', 'b', 'c']);
  });

  it('passes the 0-based row index as the second arg', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    const indices: number[] = [];
    forEachRow(ws, 'A1:A3', (_row, i) => {
      indices.push(i);
    });
    expect(indices).toEqual([0, 1]);
  });

  it('does not invoke the callback when the data area is empty', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    let calls = 0;
    forEachRow(ws, 'A1:A1', () => {
      calls++;
    });
    expect(calls).toBe(0);
  });

  it('the callback row receives the same shape as readRangeAsObjects', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    let firstRow: Record<string, unknown> | null = null;
    forEachRow(ws, 'A1:B2', (row) => {
      if (firstRow === null) firstRow = row;
    });
    expect(firstRow).toEqual({ name: 'Alice', age: 30 });
  });
});
