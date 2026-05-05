// Tests for the typed <sortState> model.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { makeSortState } from '../../src/worksheet/sort-state';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('sortState round-trip', () => {
  it('preserves a multi-condition sort with iconSet + customList + caseSensitive', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'Name');
    setCell(ws, 2, 1, 'Bob');
    setCell(ws, 3, 1, 'Alice');

    ws.sortState = makeSortState({
      ref: 'A1:C20',
      caseSensitive: true,
      sortMethod: 'pinYin',
      columnSort: false,
      conditions: [
        { ref: 'A2:A20', sortBy: 'value', descending: false, customList: 'priorityList' },
        { ref: 'B2:B20', sortBy: 'icon', iconSet: '5Arrows', iconId: 2 },
        { ref: 'C2:C20', sortBy: 'cellColor', dxfId: 0, descending: true },
      ],
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const ss = ws2.sortState;
    expect(ss).toBeDefined();
    expect(ss?.ref).toBe('A1:C20');
    expect(ss?.caseSensitive).toBe(true);
    expect(ss?.sortMethod).toBe('pinYin');
    expect(ss?.columnSort).toBe(false);
    expect(ss?.conditions.length).toBe(3);
    expect(ss?.conditions[0]?.customList).toBe('priorityList');
    expect(ss?.conditions[1]?.iconSet).toBe('5Arrows');
    expect(ss?.conditions[1]?.iconId).toBe(2);
    expect(ss?.conditions[2]?.sortBy).toBe('cellColor');
    expect(ss?.conditions[2]?.dxfId).toBe(0);
    expect(ss?.conditions[2]?.descending).toBe(true);
  });

  it('drops unknown sortMethod / sortBy / iconSet enum values', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'S');
    const ws = wb.sheets[0]?.kind === 'worksheet' ? wb.sheets[0].sheet : undefined;
    if (!ws) throw new Error('worksheet');
    setCell(ws, 1, 1, 1);
    ws.sortState = makeSortState({
      ref: 'A1:A10',
      sortMethod: 'gibberish' as never,
      conditions: [{ ref: 'A1:A10', sortBy: 'gibberish' as never, iconSet: 'gibberish' as never }],
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.sortState?.sortMethod).toBeUndefined();
    expect(ws2.sortState?.conditions[0]?.sortBy).toBeUndefined();
    expect(ws2.sortState?.conditions[0]?.iconSet).toBeUndefined();
    expect(ws2.sortState?.conditions[0]?.ref).toBe('A1:A10');
  });

  it('emits no <sortState/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'N');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.sortState).toBeUndefined();
  });
});