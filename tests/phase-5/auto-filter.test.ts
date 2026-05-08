import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { makeAutoFilter, makeFilterColumn } from '../../src/worksheet/auto-filter';
import { getAutoFilter, setAutoFilter, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('setAutoFilter / getAutoFilter', () => {
  it('starts undefined', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    expect(getAutoFilter(ws)).toBeUndefined();
  });

  it('sets, replaces, and clears the filter', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    setAutoFilter(ws, makeAutoFilter({ ref: 'A1:E10' }));
    expect(getAutoFilter(ws)?.ref).toBe('A1:E10');
    setAutoFilter(ws, makeAutoFilter({ ref: 'A1:E20' }));
    expect(getAutoFilter(ws)?.ref).toBe('A1:E20');
    setAutoFilter(ws, undefined);
    expect(getAutoFilter(ws)).toBeUndefined();
  });
});

describe('autoFilter round-trip through saveWorkbook → loadWorkbook', () => {
  it('preserves a bare ref-only filter', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    setAutoFilter(ws, makeAutoFilter({ ref: 'A1:D100' }));
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const filter = getAutoFilter(expectSheet(wb2.sheets[0]?.sheet));
    expect(filter?.ref).toBe('A1:D100');
    expect(filter?.filterColumns).toEqual([]);
  });

  it('preserves a value-list filterColumn', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    setAutoFilter(
      ws,
      makeAutoFilter({
        ref: 'A1:C10',
        filterColumns: [makeFilterColumn({ colId: 1, values: ['apple', 'banana'] })],
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const filter = getAutoFilter(expectSheet(wb2.sheets[0]?.sheet));
    expect(filter?.filterColumns.length).toBe(1);
    const fc = filter?.filterColumns[0];
    expect(fc?.colId).toBe(1);
    expect(fc?.values).toEqual(['apple', 'banana']);
  });

  it('preserves multi-column filters with blank flag', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    setAutoFilter(
      ws,
      makeAutoFilter({
        ref: 'A1:E50',
        filterColumns: [
          makeFilterColumn({ colId: 0, values: ['active'] }),
          makeFilterColumn({ colId: 3, values: ['Q1', 'Q2'], blank: true }),
        ],
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const filter = getAutoFilter(expectSheet(wb2.sheets[0]?.sheet));
    expect(filter?.filterColumns.map((fc) => fc.colId)).toEqual([0, 3]);
    expect(filter?.filterColumns[1]?.blank).toBe(true);
    expect(filter?.filterColumns[1]?.values).toEqual(['Q1', 'Q2']);
  });

  it('escapes special chars in filter values', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    setAutoFilter(
      ws,
      makeAutoFilter({
        ref: 'A1:A10',
        filterColumns: [makeFilterColumn({ colId: 0, values: ['a < b', 'c & d', '"q"'] })],
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(getAutoFilter(expectSheet(wb2.sheets[0]?.sheet))?.filterColumns[0]?.values).toEqual([
      'a < b',
      'c & d',
      '"q"',
    ]);
  });

  it('omits the <autoFilter> element when none is set', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'NoFilter');
    const bytes = await workbookToBytes(wb);
    const txt = new TextDecoder().decode(bytes);
    expect(txt).not.toContain('<autoFilter');
  });
});
