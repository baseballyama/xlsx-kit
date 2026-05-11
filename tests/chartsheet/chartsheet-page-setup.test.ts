// Tests for the typed chartsheet pageMargins / pageSetup / headerFooter.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addChartsheet, createWorkbook } from '../../src/workbook/workbook';
import { makeHeaderFooter, makePageMargins, makePageSetup } from '../../src/worksheet/page-setup';

describe('chartsheet page setup round-trip', () => {
  it('preserves pageMargins + pageSetup + headerFooter on a chartsheet', async () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    const ref = wb.sheets[0];
    if (ref?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    const cs = ref.sheet;
    cs.pageMargins = makePageMargins({ left: 0.4, right: 0.4, top: 1.1, bottom: 1.1, header: 0.3, footer: 0.3 });
    cs.pageSetup = makePageSetup({
      paperSize: 9,
      orientation: 'landscape',
      blackAndWhite: true,
      horizontalDpi: 600,
      verticalDpi: 600,
    });
    cs.headerFooter = makeHeaderFooter({
      oddHeader: '&LChartsheet&CCenter&RPage &P',
      oddFooter: '&CSecret',
      differentFirst: false,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref2 = wb2.sheets[0];
    if (ref2?.kind !== 'chartsheet') throw new Error('expected chartsheet on reload');
    const cs2 = ref2.sheet;
    expect(cs2.pageMargins?.left).toBeCloseTo(0.4);
    expect(cs2.pageMargins?.top).toBeCloseTo(1.1);
    expect(cs2.pageMargins?.header).toBeCloseTo(0.3);
    expect(cs2.pageSetup?.paperSize).toBe(9);
    expect(cs2.pageSetup?.orientation).toBe('landscape');
    expect(cs2.pageSetup?.blackAndWhite).toBe(true);
    expect(cs2.pageSetup?.horizontalDpi).toBe(600);
    expect(cs2.headerFooter?.oddHeader).toBe('&LChartsheet&CCenter&RPage &P');
    expect(cs2.headerFooter?.oddFooter).toBe('&CSecret');
    expect(cs2.headerFooter?.differentFirst).toBe(false);
  });

  it('emits no page-setup elements when undefined', async () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref = wb2.sheets[0];
    if (ref?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    expect(ref.sheet.pageMargins).toBeUndefined();
    expect(ref.sheet.pageSetup).toBeUndefined();
    expect(ref.sheet.headerFooter).toBeUndefined();
  });
});