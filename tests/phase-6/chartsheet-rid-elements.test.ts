// Round-trip tests for the chartsheet rId-link siblings:
// legacyDrawing, legacyDrawingHF, drawingHF, picture.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addChartsheet, createWorkbook } from '../../src/xlsx/workbook/workbook';

describe('chartsheet rId-link siblings round-trip', () => {
  it('preserves legacyDrawing / legacyDrawingHF / picture rIds', async () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    const ref = wb.sheets[0];
    if (ref?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    const cs = ref.sheet;
    cs.legacyDrawingRId = 'rIdLD1';
    cs.legacyDrawingHFRId = 'rIdLDHF1';
    cs.backgroundPictureRId = 'rIdPic1';

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref2 = wb2.sheets[0];
    if (ref2?.kind !== 'chartsheet') throw new Error('expected chartsheet on reload');
    expect(ref2.sheet.legacyDrawingRId).toBe('rIdLD1');
    expect(ref2.sheet.legacyDrawingHFRId).toBe('rIdLDHF1');
    expect(ref2.sheet.backgroundPictureRId).toBe('rIdPic1');
  });

  it('preserves drawingHF with per-section image indices', async () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    const ref = wb.sheets[0];
    if (ref?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    ref.sheet.drawingHF = {
      rId: 'rIdDHF1',
      lho: 1,
      cho: 2,
      rho: 3,
      lhe: 4,
      che: 5,
      rhe: 6,
      lhf: 7,
      chf: 8,
      rhf: 9,
      lfo: 10,
      cfo: 11,
      rfo: 12,
      lfe: 13,
      cfe: 14,
      rfe: 15,
      lff: 16,
      cff: 17,
      rff: 18,
    };

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref2 = wb2.sheets[0];
    if (ref2?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    const dhf = ref2.sheet.drawingHF;
    expect(dhf?.rId).toBe('rIdDHF1');
    expect(dhf?.lho).toBe(1);
    expect(dhf?.cho).toBe(2);
    expect(dhf?.rff).toBe(18);
  });

  it('emits no <legacyDrawing/>/<picture/>/<drawingHF/> when all undefined', async () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref = wb2.sheets[0];
    if (ref?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    expect(ref.sheet.legacyDrawingRId).toBeUndefined();
    expect(ref.sheet.legacyDrawingHFRId).toBeUndefined();
    expect(ref.sheet.backgroundPictureRId).toBeUndefined();
    expect(ref.sheet.drawingHF).toBeUndefined();
  });
});