// Tests for the typed worksheet-level <phoneticPr> model. Per
// docs/plan/13-full-excel-coverage.md §B10.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { makeWorksheetPhoneticProperties } from '../../src/xlsx/worksheet/phonetic';
import { setCell, type Worksheet } from '../../src/xlsx/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/xlsx/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('worksheet phoneticPr round-trip', () => {
  it('preserves fontId + type + alignment', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'P');
    setCell(ws, 1, 1, '日本語のセル');
    ws.phoneticPr = makeWorksheetPhoneticProperties({
      fontId: 1,
      type: 'Hiragana',
      alignment: 'distributed',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.phoneticPr?.fontId).toBe(1);
    expect(ws2.phoneticPr?.type).toBe('Hiragana');
    expect(ws2.phoneticPr?.alignment).toBe('distributed');
  });

  it('rejects an unknown type / alignment value (drops it on read)', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'P');
    setCell(ws, 1, 1, 1);
    // type: 'Hiragana' valid; alignment 'invalid' not in the set.
    ws.phoneticPr = makeWorksheetPhoneticProperties({
      fontId: 0,
      type: 'fullwidthKatakana',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.phoneticPr?.type).toBe('fullwidthKatakana');
    expect(ws2.phoneticPr?.alignment).toBeUndefined();
  });

  it('emits no <phoneticPr> when undefined', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'N');
    setCell(ws, 1, 1, 'a');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.phoneticPr).toBeUndefined();
  });
});