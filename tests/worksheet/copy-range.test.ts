// Tests for copyRange.

import { describe, expect, it } from 'vitest';
import { setCellFont } from '../../src/styles/cell-style';
import { makeFont } from '../../src/styles/fonts';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  copyRange,
  getRangeValues,
  setCell,
} from '../../src/worksheet/worksheet';

describe('copyRange', () => {
  it('copies cell values to a new top-left within the same sheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'c');
    setCell(ws, 2, 2, 'd');
    expect(copyRange(ws, 'A1:B2', 'D5:E6')).toBe(4);
    expect(getRangeValues(ws, 'D5:E6')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
    // Source unchanged.
    expect(getRangeValues(ws, 'A1:B2')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('preserves styleId on copied cells (shared stylesheet)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const orig = setCell(ws, 1, 1, 'x');
    setCellFont(wb, orig, makeFont({ bold: true }));
    copyRange(ws, 'A1', 'C3');
    const copy = ws.rows.get(3)?.get(3);
    if (!copy) throw new Error('expected C3');
    expect(copy.styleId).toBe(orig.styleId);
    expect(copy.value).toBe('x');
  });

  it('skips empty source cells (returns count of copied cells only)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 2, 2, 'd');
    // A1:B2 has only 2 populated; A1, B2.
    expect(copyRange(ws, 'A1:B2', 'D1:E2')).toBe(2);
    expect(ws.rows.get(1)?.get(4)?.value).toBe('a');
    expect(ws.rows.get(2)?.get(5)?.value).toBe('d');
  });

  it('truncates when target range is smaller than source', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 1, 3, 'c');
    expect(copyRange(ws, 'A1:C1', 'E1:F1')).toBe(2);
    expect(ws.rows.get(1)?.get(5)?.value).toBe('a');
    expect(ws.rows.get(1)?.get(6)?.value).toBe('b');
    expect(ws.rows.get(1)?.get(7)).toBeUndefined();
  });

  it('cross-worksheet copy via opts.targetWs', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'src');
    expect(copyRange(a, 'A1', 'C3', { targetWs: b })).toBe(1);
    expect(b.rows.get(3)?.get(3)?.value).toBe('src');
    expect(a.rows.get(3)?.get(3)).toBeUndefined();
  });

  it('overwrites existing cells in the target extent', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'src');
    setCell(ws, 5, 5, 'dst-original');
    copyRange(ws, 'A1', 'E5');
    expect(ws.rows.get(5)?.get(5)?.value).toBe('src');
  });
});
