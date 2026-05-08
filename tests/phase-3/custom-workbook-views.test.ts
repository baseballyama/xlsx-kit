// Tests for the typed workbook-level <customWorkbookViews> model.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { makeCustomWorkbookView } from '../../src/xlsx/workbook/views';

describe('customWorkbookViews round-trip', () => {
  it('preserves a saved view with full attribute set', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    wb.customWorkbookViews = [
      makeCustomWorkbookView({
        name: 'Reviewer',
        guid: '{12345678-1234-1234-1234-123456789012}',
        windowWidth: 24000,
        windowHeight: 12000,
        activeSheetId: 0,
        autoUpdate: true,
        mergeInterval: 5,
        personalView: true,
        includePrintSettings: false,
        includeHiddenRowCol: true,
        showFormulaBar: true,
        showStatusbar: true,
        showComments: 'commIndAndComment',
        showObjects: 'placeholders',
        xWindow: 100,
        yWindow: 50,
        tabRatio: 600,
      }),
    ];

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const v = wb2.customWorkbookViews?.[0];
    expect(v?.name).toBe('Reviewer');
    expect(v?.guid).toBe('{12345678-1234-1234-1234-123456789012}');
    expect(v?.windowWidth).toBe(24000);
    expect(v?.windowHeight).toBe(12000);
    expect(v?.activeSheetId).toBe(0);
    expect(v?.autoUpdate).toBe(true);
    expect(v?.mergeInterval).toBe(5);
    expect(v?.personalView).toBe(true);
    expect(v?.includePrintSettings).toBe(false);
    expect(v?.includeHiddenRowCol).toBe(true);
    expect(v?.showFormulaBar).toBe(true);
    expect(v?.showComments).toBe('commIndAndComment');
    expect(v?.showObjects).toBe('placeholders');
    expect(v?.tabRatio).toBe(600);
  });

  it('round-trips multiple entries', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    wb.customWorkbookViews = [
      makeCustomWorkbookView({ name: 'V1', guid: '{aaa}', windowWidth: 1000, windowHeight: 800, activeSheetId: 0 }),
      makeCustomWorkbookView({ name: 'V2', guid: '{bbb}', windowWidth: 1200, windowHeight: 900, activeSheetId: 1 }),
    ];

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.customWorkbookViews?.length).toBe(2);
    expect(wb2.customWorkbookViews?.[0]?.name).toBe('V1');
    expect(wb2.customWorkbookViews?.[1]?.activeSheetId).toBe(1);
  });

  it('emits no <customWorkbookViews/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.customWorkbookViews).toBeUndefined();
  });
});