// Tests for mapRange — header-driven row transform in place.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { mapRange, readRangeAsObjects, setCell } from '../../src/worksheet/worksheet';

describe('mapRange', () => {
  it('transforms each row using the callback (identity returns are no-op)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Bob'); setCell(ws, 3, 2, 25);
    mapRange(ws, 'A1:B3', (row) => {
      const age = row['age'];
      return { ...row, age: typeof age === 'number' ? age + 1 : (age ?? null) };
    });
    expect(readRangeAsObjects(ws, 'A1:B3')).toEqual([
      { name: 'Alice', age: 31 },
      { name: 'Bob', age: 26 },
    ]);
  });

  it('returning null for a key clears that cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 3, 1, 'b');
    mapRange(ws, 'A1:A3', () => ({ k: null }));
    expect(readRangeAsObjects(ws, 'A1:A3')).toEqual([{ k: null }, { k: null }]);
  });

  it('ignores extra keys in the returned row that are not in the headers', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 2, 1, 'Alice');
    mapRange(ws, 'A1:A2', (row) => ({ ...row, extra: 'ignored' }));
    // Only "name" column was in the range; the "extra" key is silently dropped.
    expect(readRangeAsObjects(ws, 'A1:A2')).toEqual([{ name: 'Alice' }]);
  });

  it('omitting a key in the returned row clears that cell to null', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'left'); setCell(ws, 2, 2, 'right');
    // Return only key 'a' — 'b' should be cleared.
    mapRange(ws, 'A1:B2', (row) => ({ a: row['a'] ?? null }));
    expect(readRangeAsObjects(ws, 'A1:B2')).toEqual([{ a: 'left', b: null }]);
  });

  it('preserves multi-column row structure (callback gets full row)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b'); setCell(ws, 1, 3, 'c');
    setCell(ws, 2, 1, 1); setCell(ws, 2, 2, 2); setCell(ws, 2, 3, 3);
    mapRange(ws, 'A1:C2', (row) => {
      const a = row['a'];
      const b = row['b'];
      const c = row['c'];
      return {
        ...row,
        c: typeof a === 'number' && typeof b === 'number' ? a + b : (c ?? null),
      };
    });
    expect(readRangeAsObjects(ws, 'A1:C2')).toEqual([{ a: 1, b: 2, c: 3 }]);
  });
});
