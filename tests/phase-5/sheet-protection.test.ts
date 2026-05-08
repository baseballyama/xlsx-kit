// Tests for the typed <sheetProtection> model. Per
// docs/plan/13-full-excel-coverage.md §B5 (without the password-hashing
// helper — saltValue/spinCount/algorithmName/hashValue round-trip
// verbatim).

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { makeSheetProtection } from '../../src/xlsx/worksheet/protection';
import { setCell, type Worksheet } from '../../src/xlsx/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/xlsx/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('sheetProtection API', () => {
  it('factory returns only the fields that were set', () => {
    const sp = makeSheetProtection({ sheet: true, formatCells: false, sort: true });
    expect(sp.sheet).toBe(true);
    expect(sp.formatCells).toBe(false);
    expect(sp.sort).toBe(true);
    expect(sp.objects).toBeUndefined();
    expect(sp.saltValue).toBeUndefined();
  });

  it('round-trips all 16 boolean lock flags', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'L');
    setCell(ws, 1, 1, 1);
    ws.sheetProtection = makeSheetProtection({
      sheet: true,
      objects: true,
      scenarios: true,
      formatCells: false,
      formatColumns: false,
      formatRows: false,
      insertColumns: false,
      insertRows: false,
      insertHyperlinks: false,
      deleteColumns: false,
      deleteRows: false,
      selectLockedCells: false,
      selectUnlockedCells: true,
      sort: false,
      autoFilter: false,
      pivotTables: false,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const sp = ws2.sheetProtection;
    expect(sp).toBeDefined();
    expect(sp?.sheet).toBe(true);
    expect(sp?.objects).toBe(true);
    expect(sp?.scenarios).toBe(true);
    expect(sp?.formatCells).toBe(false);
    expect(sp?.selectUnlockedCells).toBe(true);
    expect(sp?.pivotTables).toBe(false);
  });

  it('round-trips password-hash fields verbatim', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'P');
    setCell(ws, 1, 1, 1);
    ws.sheetProtection = makeSheetProtection({
      sheet: true,
      algorithmName: 'SHA-512',
      hashValue: 'abc123==',
      saltValue: 'def456==',
      spinCount: 100000,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const sp = ws2.sheetProtection;
    expect(sp?.algorithmName).toBe('SHA-512');
    expect(sp?.hashValue).toBe('abc123==');
    expect(sp?.saltValue).toBe('def456==');
    expect(sp?.spinCount).toBe(100000);
  });

  it('emits no <sheetProtection/> when all fields undefined', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'N');
    setCell(ws, 1, 1, 'a');
    ws.sheetProtection = makeSheetProtection({});

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.sheetProtection).toBeUndefined();
  });
});