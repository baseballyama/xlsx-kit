// Tests for chartsheet-level <customSheetViews>.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addChartsheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { makeChartsheetCustomSheetView } from '../../src/xlsx/chartsheet/chartsheet';
import { makeHeaderFooter, makePageMargins, makePageSetup } from '../../src/xlsx/worksheet/page-setup';

describe('chartsheet customSheetViews round-trip', () => {
  it('preserves a saved view with scale + state + nested page setup', async () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    const ref = wb.sheets[0];
    if (ref?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    ref.sheet.customSheetViews.push(
      makeChartsheetCustomSheetView({
        guid: '{aabbccdd-0000-1111-2222-333344445555}',
        scale: 75,
        state: 'hidden',
        zoomToFit: true,
        pageMargins: makePageMargins({ left: 0.4, right: 0.4, top: 1, bottom: 1, header: 0.3, footer: 0.3 }),
        pageSetup: makePageSetup({ paperSize: 9, orientation: 'portrait' }),
        headerFooter: makeHeaderFooter({ oddHeader: '&CSaved view' }),
      }),
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref2 = wb2.sheets[0];
    if (ref2?.kind !== 'chartsheet') throw new Error('expected chartsheet on reload');
    expect(ref2.sheet.customSheetViews.length).toBe(1);
    const v = ref2.sheet.customSheetViews[0];
    expect(v?.guid).toBe('{aabbccdd-0000-1111-2222-333344445555}');
    expect(v?.scale).toBe(75);
    expect(v?.state).toBe('hidden');
    expect(v?.zoomToFit).toBe(true);
    expect(v?.pageMargins?.left).toBeCloseTo(0.4);
    expect(v?.pageSetup?.paperSize).toBe(9);
    expect(v?.pageSetup?.orientation).toBe('portrait');
    expect(v?.headerFooter?.oddHeader).toBe('&CSaved view');
  });

  it('drops unknown state enum values', async () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    const ref = wb.sheets[0];
    if (ref?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    ref.sheet.customSheetViews.push(
      makeChartsheetCustomSheetView({ guid: '{xx}', state: 'gibberish' as never }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref2 = wb2.sheets[0];
    if (ref2?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    expect(ref2.sheet.customSheetViews[0]?.state).toBeUndefined();
    expect(ref2.sheet.customSheetViews[0]?.guid).toBe('{xx}');
  });

  it('emits no <customSheetViews/> when empty', async () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref = wb2.sheets[0];
    if (ref?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    expect(ref.sheet.customSheetViews.length).toBe(0);
  });
});