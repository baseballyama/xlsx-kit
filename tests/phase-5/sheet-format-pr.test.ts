// Tests for the extended <sheetFormatPr> attributes (B11). Per
// docs/plan/13-full-excel-coverage.md §B11.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  setCell,
  setColumnDimension,
  setRowDimension,
  type Worksheet,
} from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('sheetFormatPr round-trip', () => {
  it('outlineLevelRow + outlineLevelCol auto-compute from row/column dimensions', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'OL');
    setCell(ws, 1, 1, 1);
    setRowDimension(ws, 2, { outlineLevel: 1 });
    setRowDimension(ws, 3, { outlineLevel: 2 });
    setRowDimension(ws, 4, { outlineLevel: 1 });
    setColumnDimension(ws, 3, { outlineLevel: 2 });
    expect(ws.outlineLevelRow).toBeUndefined();
    expect(ws.outlineLevelCol).toBeUndefined();

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    // After round-trip the explicit value comes back from <sheetFormatPr>.
    expect(ws2.outlineLevelRow).toBe(2);
    expect(ws2.outlineLevelCol).toBe(2);
  });

  it('explicit outlineLevelRow overrides auto-compute', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'OL2');
    setCell(ws, 1, 1, 1);
    setRowDimension(ws, 2, { outlineLevel: 1 });
    ws.outlineLevelRow = 7;

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.outlineLevelRow).toBe(7);
  });

  it('round-trips customHeight / zeroHeight / thickTop / thickBottom / baseColWidth', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    setCell(ws, 1, 1, 1);
    ws.customHeight = true;
    ws.zeroHeight = false;
    ws.thickTop = true;
    ws.thickBottom = false;
    ws.baseColWidth = 10;

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.customHeight).toBe(true);
    expect(ws2.zeroHeight).toBe(false);
    expect(ws2.thickTop).toBe(true);
    expect(ws2.thickBottom).toBe(false);
    expect(ws2.baseColWidth).toBe(10);
  });

  it('emits no <sheetFormatPr> when nothing is set', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Plain');
    setCell(ws, 1, 1, 'a');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.outlineLevelRow).toBeUndefined();
    expect(ws2.outlineLevelCol).toBeUndefined();
    expect(ws2.customHeight).toBeUndefined();
    expect(ws2.defaultColumnWidth).toBeUndefined();
  });
});