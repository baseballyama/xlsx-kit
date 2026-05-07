// Tests for writeRange — write a 2D values block at an arbitrary anchor.

import { describe, expect, it } from 'vitest';
import { setBold } from '../../src/styles/cell-style';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell, writeRange } from '../../src/worksheet/worksheet';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('writeRange', () => {
  it('writes a 3×2 block starting at B2', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const result = writeRange(ws, 'B2', [
      ['a', 'b'],
      ['c', 'd'],
      ['e', 'f'],
    ]);
    expect(result).toEqual({ minRow: 2, maxRow: 4, minCol: 2, maxCol: 3 });
    expect(cellAt(ws, 2, 2).value).toBe('a');
    expect(cellAt(ws, 2, 3).value).toBe('b');
    expect(cellAt(ws, 3, 2).value).toBe('c');
    expect(cellAt(ws, 4, 3).value).toBe('f');
  });

  it('skips null / undefined entries (no cell created)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    writeRange(ws, 'A1', [
      ['x', null, 'z'],
      [undefined, 'y', undefined],
    ]);
    expect(ws.rows.get(1)?.has(2)).toBe(false);
    expect(cellAt(ws, 1, 1).value).toBe('x');
    expect(cellAt(ws, 1, 3).value).toBe('z');
    expect(ws.rows.get(2)?.has(1)).toBe(false);
    expect(cellAt(ws, 2, 2).value).toBe('y');
    expect(ws.rows.get(2)?.has(3)).toBe(false);
  });

  it('returns undefined for an empty input', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(writeRange(ws, 'A1', [])).toBeUndefined();
    expect(ws.rows.size).toBe(0);
  });

  it('overwrites pre-existing cells in place (styleId preserved)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'old');
    setBold(wb, cellAt(ws, 1, 1));
    const styledId = cellAt(ws, 1, 1).styleId;
    expect(styledId).not.toBe(0);
    writeRange(ws, 'A1', [['new']]);
    expect(cellAt(ws, 1, 1).value).toBe('new');
    // setCell preserves the existing cell instance + its styleId.
    expect(cellAt(ws, 1, 1).styleId).toBe(styledId);
  });

  it('handles a single-cell write', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const result = writeRange(ws, 'D7', [['solo']]);
    expect(result).toEqual({ minRow: 7, maxRow: 7, minCol: 4, maxCol: 4 });
    expect(cellAt(ws, 7, 4).value).toBe('solo');
  });
});
