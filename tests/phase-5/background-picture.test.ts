// Tests for the typed worksheet <picture> background-image rId.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('worksheet <picture> background round-trip', () => {
  it('preserves backgroundPictureRId', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'P');
    setCell(ws, 1, 1, 1);
    ws.backgroundPictureRId = 'rIdBg1';

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.backgroundPictureRId).toBe('rIdBg1');
  });

  it('emits no <picture/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'N');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.backgroundPictureRId).toBeUndefined();
  });
});