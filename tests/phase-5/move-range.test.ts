// Tests for moveRange.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  getRangeValues,
  moveRange,
  setCell,
} from '../../src/worksheet/worksheet';

describe('moveRange', () => {
  it('moves cells to a non-overlapping target and clears the source', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'c');
    expect(moveRange(ws, 'A1:B2', 'D5:E6')).toBe(3);
    expect(ws.rows.get(1)).toBeUndefined();
    expect(ws.rows.get(2)).toBeUndefined();
    expect(getRangeValues(ws, 'D5:E6')).toEqual([
      ['a', 'b'],
      ['c', null],
    ]);
  });

  it('overlapping move keeps the destination intact', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 1, 3, 'c');
    // Shift A1:C1 → B1:D1 (overlap = B1, C1).
    moveRange(ws, 'A1:C1', 'B1:D1');
    expect(ws.rows.get(1)?.get(1)).toBeUndefined();
    expect(ws.rows.get(1)?.get(2)?.value).toBe('a');
    expect(ws.rows.get(1)?.get(3)?.value).toBe('b');
    expect(ws.rows.get(1)?.get(4)?.value).toBe('c');
  });

  it('cross-worksheet move via opts.targetWs clears source on the source sheet', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'src');
    expect(moveRange(a, 'A1', 'C3', { targetWs: b })).toBe(1);
    expect(a.rows.get(1)).toBeUndefined();
    expect(b.rows.get(3)?.get(3)?.value).toBe('src');
  });

  it('returns 0 for empty source range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(moveRange(ws, 'A1:B2', 'D5:E6')).toBe(0);
  });
});
