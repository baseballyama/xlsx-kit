// Tests for writeRangeFromObjects — write Record[] starting at A1 anchor.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  readRangeAsObjects,
  writeRangeFromObjects,
} from '../../src/worksheet/worksheet';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('writeRangeFromObjects', () => {
  it('writes header + rows starting at the anchor', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const result = writeRangeFromObjects(ws, 'A1', [
      { name: 'Alice', age: 30 },
      { name: 'Bob', age: 25 },
    ]);
    expect(result).toEqual({ minRow: 1, maxRow: 3, minCol: 1, maxCol: 2 });
    expect(cellAt(ws, 1, 1).value).toBe('name');
    expect(cellAt(ws, 1, 2).value).toBe('age');
    expect(cellAt(ws, 2, 1).value).toBe('Alice');
    expect(cellAt(ws, 3, 2).value).toBe(25);
  });

  it('returns undefined for empty input', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(writeRangeFromObjects(ws, 'A1', [])).toBeUndefined();
    expect(ws.rows.size).toBe(0);
  });

  it('opts.headers pins column order regardless of object key order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    writeRangeFromObjects(
      ws,
      'A1',
      [{ b: 2, a: 1 }],
      { headers: ['a', 'b'] },
    );
    expect(cellAt(ws, 1, 1).value).toBe('a');
    expect(cellAt(ws, 1, 2).value).toBe('b');
    expect(cellAt(ws, 2, 1).value).toBe(1);
    expect(cellAt(ws, 2, 2).value).toBe(2);
  });

  it('null / undefined values skip the cell (not even materialised)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    writeRangeFromObjects(ws, 'A1', [
      { x: 'a', y: null, z: 'c' },
      { x: undefined, y: 'b', z: undefined },
    ]);
    expect(ws.rows.get(2)?.has(2)).toBe(false);
    expect(cellAt(ws, 2, 1).value).toBe('a');
    expect(cellAt(ws, 2, 3).value).toBe('c');
    expect(ws.rows.get(3)?.has(1)).toBe(false);
    expect(cellAt(ws, 3, 2).value).toBe('b');
  });

  it('headers default to union of keys (first-appearance order)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    writeRangeFromObjects(ws, 'A1', [
      { a: 1, b: 2 },
      { a: 3, c: 4 }, // new key 'c' appears after 'b' in header order
    ]);
    expect(cellAt(ws, 1, 1).value).toBe('a');
    expect(cellAt(ws, 1, 2).value).toBe('b');
    expect(cellAt(ws, 1, 3).value).toBe('c');
  });

  it('round-trips through readRangeAsObjects', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const data = [
      { id: 1, name: 'a' },
      { id: 2, name: 'b' },
      { id: 3, name: 'c' },
    ];
    const r = writeRangeFromObjects(ws, 'C5', data);
    if (!r) throw new Error('write returned undefined');
    const range = `C${r.minRow}:D${r.maxRow}`;
    expect(readRangeAsObjects(ws, range)).toEqual(data);
  });
});
