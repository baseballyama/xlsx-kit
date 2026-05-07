// Tests for addTableFromObjects — write Record[] + register Excel table.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { addTableFromObjects } from '../../src/worksheet/table';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('addTableFromObjects', () => {
  it('writes the data and registers a table over the bounding-box', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const table = addTableFromObjects(wb, ws, {
      name: 'People',
      startRef: 'A1',
      objects: [
        { name: 'Alice', age: 30 },
        { name: 'Bob', age: 25 },
      ],
    });
    expect(table.name).toBe('People');
    expect(table.displayName).toBe('People');
    expect(table.ref).toBe('A1:B3');
    expect(table.columns.map((c) => c.name)).toEqual(['name', 'age']);
    // Cells are also written
    expect(cellAt(ws, 1, 1).value).toBe('name');
    expect(cellAt(ws, 2, 1).value).toBe('Alice');
    expect(cellAt(ws, 3, 2).value).toBe(25);
    // Table is registered on the worksheet
    expect(ws.tables.length).toBe(1);
    expect(ws.tables[0]).toBe(table);
  });

  it('opts.headers pins column order in both the data write and the table columns', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const table = addTableFromObjects(wb, ws, {
      name: 'T',
      startRef: 'A1',
      objects: [{ b: 2, a: 1 }],
      headers: ['a', 'b'],
    });
    expect(table.columns.map((c) => c.name)).toEqual(['a', 'b']);
    expect(cellAt(ws, 1, 1).value).toBe('a');
    expect(cellAt(ws, 1, 2).value).toBe('b');
  });

  it('throws on empty objects (no zero-row tables)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => addTableFromObjects(wb, ws, { name: 'T', startRef: 'A1', objects: [] })).toThrow(/non-empty/);
  });

  it('opts.style sets a built-in style on the registered table', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const table = addTableFromObjects(wb, ws, {
      name: 'T',
      startRef: 'A1',
      objects: [{ a: 1 }],
      style: 'TableStyleMedium2',
    });
    expect(table.styleInfo?.name).toBe('TableStyleMedium2');
  });
});
