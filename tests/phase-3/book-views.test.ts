// Tests for the typed workbook-level <bookViews> model.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { makeWorkbookView } from '../../src/workbook/views';
import { setCell } from '../../src/worksheet/worksheet';

describe('bookViews round-trip', () => {
  it('preserves a single workbookView with firstSheet / activeTab / window position', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    const ws = wb.sheets[2]?.kind === 'worksheet' ? wb.sheets[2].sheet : undefined;
    if (ws) setCell(ws, 1, 1, 'on C');

    wb.bookViews = [
      makeWorkbookView({
        firstSheet: 1,
        activeTab: 2,
        xWindow: 240,
        yWindow: 60,
        windowWidth: 24000,
        windowHeight: 12000,
        tabRatio: 700,
        showSheetTabs: true,
      }),
    ];

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.bookViews?.length).toBe(1);
    const v = wb2.bookViews?.[0];
    expect(v?.firstSheet).toBe(1);
    expect(v?.activeTab).toBe(2);
    expect(v?.xWindow).toBe(240);
    expect(v?.yWindow).toBe(60);
    expect(v?.windowWidth).toBe(24000);
    expect(v?.windowHeight).toBe(12000);
    expect(v?.tabRatio).toBe(700);
    expect(v?.showSheetTabs).toBe(true);
  });

  it('round-trips visibility = "hidden" + the rare flags', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    wb.bookViews = [
      makeWorkbookView({
        visibility: 'hidden',
        minimized: true,
        autoFilterDateGrouping: false,
        showHorizontalScroll: false,
        showVerticalScroll: false,
      }),
    ];

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const v = wb2.bookViews?.[0];
    expect(v?.visibility).toBe('hidden');
    expect(v?.minimized).toBe(true);
    expect(v?.autoFilterDateGrouping).toBe(false);
    expect(v?.showHorizontalScroll).toBe(false);
    expect(v?.showVerticalScroll).toBe(false);
  });

  it('round-trips multiple workbookView entries', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    wb.bookViews = [makeWorkbookView({ activeTab: 0 }), makeWorkbookView({ activeTab: 0, minimized: true })];

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.bookViews?.length).toBe(2);
    expect(wb2.bookViews?.[1]?.minimized).toBe(true);
  });

  it('emits no <bookViews/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.bookViews).toBeUndefined();
  });
});