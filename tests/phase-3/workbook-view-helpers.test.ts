// Tests for the get/set helpers over Workbook.bookViews[0].

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import {
  getActiveTab,
  getFirstSheet,
  setActiveTab,
  setFirstSheet,
  setShowSheetTabs,
  setTabRatio,
  setWorkbookWindow,
} from '../../src/xlsx/workbook/views';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';

describe('workbook view helpers', () => {
  it('getActiveTab returns 0 on a fresh workbook with no bookViews', () => {
    const wb = createWorkbook();
    expect(wb.bookViews).toBeUndefined();
    expect(getActiveTab(wb)).toBe(0);
    expect(getFirstSheet(wb)).toBe(0);
  });

  it('setActiveTab + setFirstSheet allocate the primary view lazily', () => {
    const wb = createWorkbook();
    setActiveTab(wb, 2);
    expect(wb.bookViews?.length).toBe(1);
    expect(wb.bookViews?.[0]?.activeTab).toBe(2);

    setFirstSheet(wb, 1);
    expect(wb.bookViews?.[0]?.firstSheet).toBe(1);
    // Doesn't add a second view.
    expect(wb.bookViews?.length).toBe(1);
  });

  it('setTabRatio + setShowSheetTabs + setWorkbookWindow round-trip', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    setActiveTab(wb, 2);
    setFirstSheet(wb, 1);
    setTabRatio(wb, 750);
    setShowSheetTabs(wb, true);
    setWorkbookWindow(wb, { xWindow: 200, yWindow: 100, windowWidth: 25000, windowHeight: 12500 });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(getActiveTab(wb2)).toBe(2);
    expect(getFirstSheet(wb2)).toBe(1);
    expect(wb2.bookViews?.[0]?.tabRatio).toBe(750);
    expect(wb2.bookViews?.[0]?.showSheetTabs).toBe(true);
    expect(wb2.bookViews?.[0]?.xWindow).toBe(200);
    expect(wb2.bookViews?.[0]?.yWindow).toBe(100);
    expect(wb2.bookViews?.[0]?.windowWidth).toBe(25000);
    expect(wb2.bookViews?.[0]?.windowHeight).toBe(12500);
  });

  it('setWorkbookWindow only touches axes the caller supplies', () => {
    const wb = createWorkbook();
    setWorkbookWindow(wb, { xWindow: 50, yWindow: 60 });
    expect(wb.bookViews?.[0]?.xWindow).toBe(50);
    expect(wb.bookViews?.[0]?.yWindow).toBe(60);
    expect(wb.bookViews?.[0]?.windowWidth).toBeUndefined();

    setWorkbookWindow(wb, { windowWidth: 9000 });
    expect(wb.bookViews?.[0]?.xWindow).toBe(50); // unchanged
    expect(wb.bookViews?.[0]?.windowWidth).toBe(9000);
  });
});