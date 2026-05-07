// Tests for fillColumn — bulk single-column data write (header preserved).

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { fillColumn, readRangeAsObjects, setCell } from '../../src/worksheet/worksheet';

describe('fillColumn', () => {
  it('stamps the same value into every data row of the column', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 1, 2, 'flag');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    setCell(ws, 4, 1, 'c');
    fillColumn(ws, 'A1:B4', 'flag', true);
    expect(readRangeAsObjects(ws, 'A1:B4')).toEqual([
      { k: 'a', flag: true },
      { k: 'b', flag: true },
      { k: 'c', flag: true },
    ]);
  });

  it('computes per-row values via a function (row + index)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 1, 2, 'rank');
    setCell(ws, 2, 1, 'Alice');
    setCell(ws, 3, 1, 'Bob');
    setCell(ws, 4, 1, 'Carol');
    fillColumn(ws, 'A1:B4', 'rank', (_row, i) => i + 1);
    expect(readRangeAsObjects(ws, 'A1:B4')).toEqual([
      { name: 'Alice', rank: 1 },
      { name: 'Bob', rank: 2 },
      { name: 'Carol', rank: 3 },
    ]);
  });

  it('preserves the header cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'value');
    setCell(ws, 2, 1, 'old');
    fillColumn(ws, 'A1:A2', 'value', 'new');
    expect(ws.rows.get(1)?.get(1)?.value).toBe('value');
    expect(ws.rows.get(2)?.get(1)?.value).toBe('new');
  });

  it('throws when the column is not in the header row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    expect(() => fillColumn(ws, 'A1:A2', 'missing', 'x')).toThrow(/missing/);
  });

  it('no-op when the data area is empty (header-only range)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'value');
    fillColumn(ws, 'A1:A1', 'value', 'x'); // does not throw, no data rows to fill
    expect(ws.rows.get(1)?.get(1)?.value).toBe('value');
  });
});
