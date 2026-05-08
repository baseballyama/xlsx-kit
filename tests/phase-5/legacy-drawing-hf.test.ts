// Tests for the typed worksheet <legacyDrawingHF> rId.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { setCell, type Worksheet } from '../../src/xlsx/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/xlsx/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('worksheet <legacyDrawingHF> round-trip', () => {
  it('preserves the rId', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'L');
    setCell(ws, 1, 1, 1);
    ws.legacyDrawingHFRId = 'rIdHF1';

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.legacyDrawingHFRId).toBe('rIdHF1');
  });

  it('emits no <legacyDrawingHF/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'N');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.legacyDrawingHFRId).toBeUndefined();
  });
});