// Round-trip tests for the worksheet-level <customSheetViews> model.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { makeCustomSheetView } from '../../src/xlsx/worksheet/custom-sheet-views';
import { makeHeaderFooter, makePageMargins, makePageSetup, makePrintOptions } from '../../src/xlsx/worksheet/page-setup';
import { setCell, type Worksheet } from '../../src/xlsx/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/xlsx/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('worksheet customSheetViews round-trip', () => {
  it('preserves a saved view with attrs + nested page setup + breaks', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    setCell(ws, 1, 1, 'a');
    ws.customSheetViews.push(
      makeCustomSheetView({
        guid: '{12345678-1111-2222-3333-444455556666}',
        scale: 75,
        colorId: 17,
        showPageBreaks: false,
        showFormulas: false,
        showGridLines: true,
        showRowCol: true,
        outlineSymbols: true,
        zeroValues: true,
        fitToPage: false,
        printArea: false,
        filter: false,
        showAutoFilter: true,
        hiddenRows: false,
        hiddenColumns: false,
        state: 'visible',
        filterUnique: false,
        view: 'pageBreakPreview',
        showRuler: true,
        topLeftCell: 'A1',
        rowBreaks: [{ id: 5, max: 16383, man: true }],
        colBreaks: [{ id: 4, max: 1048575, man: true }],
        pageMargins: makePageMargins({ left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.3, footer: 0.3 }),
        printOptions: makePrintOptions({ horizontalCentered: true, gridLines: true }),
        pageSetup: makePageSetup({ paperSize: 9, orientation: 'landscape' }),
        headerFooter: makeHeaderFooter({ oddHeader: '&CSaved view' }),
      }),
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.customSheetViews.length).toBe(1);
    const v = ws2.customSheetViews[0];
    expect(v?.guid).toBe('{12345678-1111-2222-3333-444455556666}');
    expect(v?.scale).toBe(75);
    expect(v?.colorId).toBe(17);
    expect(v?.showPageBreaks).toBe(false);
    expect(v?.showAutoFilter).toBe(true);
    expect(v?.state).toBe('visible');
    expect(v?.view).toBe('pageBreakPreview');
    expect(v?.topLeftCell).toBe('A1');
    expect(v?.rowBreaks?.length).toBe(1);
    expect(v?.rowBreaks?.[0]?.id).toBe(5);
    expect(v?.colBreaks?.length).toBe(1);
    expect(v?.pageMargins?.left).toBeCloseTo(0.5);
    expect(v?.printOptions?.horizontalCentered).toBe(true);
    expect(v?.pageSetup?.paperSize).toBe(9);
    expect(v?.headerFooter?.oddHeader).toBe('&CSaved view');
  });

  it('drops unknown state / view enum values', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'X');
    setCell(ws, 1, 1, 1);
    ws.customSheetViews.push(
      makeCustomSheetView({
        guid: '{aa}',
        state: 'gibberish' as never,
        view: 'gibberish' as never,
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.customSheetViews[0]?.state).toBeUndefined();
    expect(ws2.customSheetViews[0]?.view).toBeUndefined();
    expect(ws2.customSheetViews[0]?.guid).toBe('{aa}');
  });

  it('emits no <customSheetViews/> when empty', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'N');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.customSheetViews.length).toBe(0);
  });
});