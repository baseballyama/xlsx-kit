// Tests for the typed <protectedRanges> model. Per
// docs/plan/13-full-excel-coverage.md §B5 (sheet-protection family).

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { parseMultiCellRange } from '../../src/xlsx/worksheet/cell-range';
import { makeProtectedRange } from '../../src/xlsx/worksheet/protected-ranges';
import { setCell, type Worksheet } from '../../src/xlsx/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/xlsx/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('protectedRanges round-trip', () => {
  it('preserves multiple ranges with different password styles', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'P');
    setCell(ws, 1, 1, 1);
    ws.protectedRanges.push(
      makeProtectedRange({
        sqref: parseMultiCellRange('A1:B5'),
        name: 'Editor1',
        password: 'CC1A',
      }),
    );
    ws.protectedRanges.push(
      makeProtectedRange({
        sqref: parseMultiCellRange('D1:E10 G1:G5'),
        name: 'Editor2',
        algorithmName: 'SHA-512',
        hashValue: 'aGFzaA==',
        saltValue: 'c2FsdA==',
        spinCount: 100000,
      }),
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.protectedRanges.length).toBe(2);
    expect(ws2.protectedRanges[0]?.name).toBe('Editor1');
    expect(ws2.protectedRanges[0]?.password).toBe('CC1A');
    expect(ws2.protectedRanges[1]?.name).toBe('Editor2');
    expect(ws2.protectedRanges[1]?.algorithmName).toBe('SHA-512');
    expect(ws2.protectedRanges[1]?.hashValue).toBe('aGFzaA==');
    expect(ws2.protectedRanges[1]?.spinCount).toBe(100000);
    expect(ws2.protectedRanges[1]?.sqref.ranges.length).toBe(2);
  });

  it('emits no <protectedRanges/> when empty', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'N');
    setCell(ws, 1, 1, 'a');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.protectedRanges.length).toBe(0);
  });
});