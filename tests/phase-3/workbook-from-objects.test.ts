// Tests for createWorkbookFromObjects — one-shot Record[] → Workbook constructor.

import { describe, expect, it } from 'vitest';
import { createWorkbookFromObjects } from '../../src/workbook/workbook';
import { getCellByCoord, listTables } from '../../src/worksheet/worksheet';

describe('createWorkbookFromObjects', () => {
  it('creates Sheet1 with header + data rows when asTable is omitted', () => {
    const wb = createWorkbookFromObjects([
      { name: 'Alice', age: 30 },
      { name: 'Bob', age: 25 },
    ]);
    expect(wb.sheets.length).toBe(1);
    const ref = wb.sheets[0];
    if (!ref || ref.kind !== 'worksheet') throw new Error('expected a worksheet ref');
    expect(ref.sheet.title).toBe('Sheet1');
    expect(getCellByCoord(ref.sheet, 'A1')?.value).toBe('name');
    expect(getCellByCoord(ref.sheet, 'A2')?.value).toBe('Alice');
    expect(getCellByCoord(ref.sheet, 'B3')?.value).toBe(25);
    // No table registered when asTable is false / omitted
    expect(listTables(ref.sheet).length).toBe(0);
  });

  it('honours opts.sheetTitle', () => {
    const wb = createWorkbookFromObjects([{ x: 1 }], { sheetTitle: 'People' });
    expect(wb.sheets[0]?.sheet.title).toBe('People');
  });

  it('asTable: true registers an Excel Table over the data', () => {
    const wb = createWorkbookFromObjects(
      [
        { id: 1, name: 'a' },
        { id: 2, name: 'b' },
      ],
      { asTable: true, tableName: 'People', style: 'TableStyleMedium2' },
    );
    const ref = wb.sheets[0];
    if (!ref || ref.kind !== 'worksheet') throw new Error('expected a worksheet ref');
    const tables = listTables(ref.sheet);
    expect(tables.length).toBe(1);
    const t = tables[0];
    if (!t) throw new Error('table missing');
    expect(t.name).toBe('People');
    expect(t.ref).toBe('A1:B3');
    expect(t.styleInfo?.name).toBe('TableStyleMedium2');
  });

  it('opts.headers pins column order', () => {
    const wb = createWorkbookFromObjects([{ b: 2, a: 1 }], { headers: ['a', 'b'] });
    const ref = wb.sheets[0];
    if (!ref || ref.kind !== 'worksheet') throw new Error('expected a worksheet ref');
    expect(getCellByCoord(ref.sheet, 'A1')?.value).toBe('a');
    expect(getCellByCoord(ref.sheet, 'B1')?.value).toBe('b');
  });

  it('returns a workbook with an empty sheet for [] input (no throw, no table)', () => {
    const wb = createWorkbookFromObjects([]);
    const ref = wb.sheets[0];
    if (!ref || ref.kind !== 'worksheet') throw new Error('expected a worksheet ref');
    expect(ref.sheet.rows.size).toBe(0);
    expect(listTables(ref.sheet).length).toBe(0);
  });
});
