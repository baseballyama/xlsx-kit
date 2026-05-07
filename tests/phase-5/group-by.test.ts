// Tests for groupBy — header-driven row grouping.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { groupBy, setCell } from '../../src/worksheet/worksheet';

describe('groupBy', () => {
  it('groups rows by the value of the chosen column', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 1, 2, 'team');
    setCell(ws, 2, 1, 'Alice');
    setCell(ws, 2, 2, 'red');
    setCell(ws, 3, 1, 'Bob');
    setCell(ws, 3, 2, 'blue');
    setCell(ws, 4, 1, 'Carol');
    setCell(ws, 4, 2, 'red');
    expect(groupBy(ws, 'A1:B4', 'team')).toEqual({
      red: [
        { name: 'Alice', team: 'red' },
        { name: 'Carol', team: 'red' },
      ],
      blue: [{ name: 'Bob', team: 'blue' }],
    });
  });

  it('returns a single bucket when every row has the same key', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 1, 2, 'v');
    setCell(ws, 2, 1, 'x');
    setCell(ws, 2, 2, 1);
    setCell(ws, 3, 1, 'x');
    setCell(ws, 3, 2, 2);
    expect(groupBy(ws, 'A1:B3', 'k')).toEqual({
      x: [
        { k: 'x', v: 1 },
        { k: 'x', v: 2 },
      ],
    });
  });

  it('null group keys collapse to the empty-string bucket', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 1, 2, 'v');
    setCell(ws, 2, 1, 'x');
    setCell(ws, 2, 2, 1);
    // row 3 col 1 is intentionally empty → key becomes ""
    setCell(ws, 3, 2, 2);
    const r = groupBy(ws, 'A1:B3', 'k');
    expect(Object.keys(r).sort()).toEqual(['', 'x']);
    expect(r['']?.[0]?.['v']).toBe(2);
  });

  it('throws when byColumn is not one of the range headers', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 2, 1, 'Alice');
    expect(() => groupBy(ws, 'A1:A2', 'team')).toThrow(/team/);
  });

  it('returns {} when the range covers only the header row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    expect(groupBy(ws, 'A1:A1', 'k')).toEqual({});
  });
});
