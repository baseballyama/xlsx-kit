// Tests for the typed <dataConsolidate> model. Per
// docs/plan/13-full-excel-coverage.md §B7 (sheet view 拡張) sub-piece.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { makeDataConsolidate } from '../../src/worksheet/data-consolidate';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('worksheet dataConsolidate round-trip', () => {
  it('preserves function + label flags + dataRefs', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    setCell(ws, 1, 1, 1);
    ws.dataConsolidate = makeDataConsolidate({
      function: 'sum',
      topLabels: true,
      leftLabels: false,
      link: true,
      dataRefs: [
        { ref: "'North'!$A$1:$D$10", sheet: 'North' },
        { ref: "'South'!$A$1:$D$10", sheet: 'South', name: 'southBlock' },
      ],
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const dc = ws2.dataConsolidate;
    expect(dc?.function).toBe('sum');
    expect(dc?.topLabels).toBe(true);
    expect(dc?.leftLabels).toBe(false);
    expect(dc?.link).toBe(true);
    expect(dc?.dataRefs?.length).toBe(2);
    expect(dc?.dataRefs?.[0]?.ref).toBe("'North'!$A$1:$D$10");
    expect(dc?.dataRefs?.[1]?.name).toBe('southBlock');
  });

  it('handles a function-only entry with no dataRefs (self-closing emit)', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    setCell(ws, 1, 1, 1);
    ws.dataConsolidate = makeDataConsolidate({ function: 'average' });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.dataConsolidate?.function).toBe('average');
    expect(ws2.dataConsolidate?.dataRefs).toBeUndefined();
  });

  it('emits no <dataConsolidate> when undefined', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'N');
    setCell(ws, 1, 1, 'a');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.dataConsolidate).toBeUndefined();
  });
});