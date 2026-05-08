// Tests for workbook-view scroll-bar / minimised / visibility toggles.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import {
  setShowHorizontalScroll,
  setShowVerticalScroll,
  setWorkbookMinimized,
  setWorkbookVisibility,
} from '../../src/workbook/views';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

describe('setShowHorizontalScroll / setShowVerticalScroll', () => {
  it('toggle independently on the primary workbookView', () => {
    const wb = createWorkbook();
    setShowHorizontalScroll(wb, false);
    expect(wb.bookViews?.[0]?.showHorizontalScroll).toBe(false);
    expect(wb.bookViews?.[0]?.showVerticalScroll).toBeUndefined();
    setShowVerticalScroll(wb, false);
    expect(wb.bookViews?.[0]?.showVerticalScroll).toBe(false);
  });

  it('lazily creates the primary workbookView', () => {
    const wb = createWorkbook();
    expect(wb.bookViews).toBeUndefined();
    setShowHorizontalScroll(wb, true);
    expect(wb.bookViews?.length).toBe(1);
  });
});

describe('setWorkbookMinimized', () => {
  it('flips the minimized flag', () => {
    const wb = createWorkbook();
    setWorkbookMinimized(wb, true);
    expect(wb.bookViews?.[0]?.minimized).toBe(true);
    setWorkbookMinimized(wb, false);
    expect(wb.bookViews?.[0]?.minimized).toBe(false);
  });
});

describe('setWorkbookVisibility', () => {
  it('accepts visible / hidden / veryHidden', () => {
    const wb = createWorkbook();
    setWorkbookVisibility(wb, 'hidden');
    expect(wb.bookViews?.[0]?.visibility).toBe('hidden');
    setWorkbookVisibility(wb, 'veryHidden');
    expect(wb.bookViews?.[0]?.visibility).toBe('veryHidden');
    setWorkbookVisibility(wb, 'visible');
    expect(wb.bookViews?.[0]?.visibility).toBe('visible');
  });
});

describe('view-toggle round-trip', () => {
  it('all four toggles survive saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setShowHorizontalScroll(wb, false);
    setShowVerticalScroll(wb, false);
    setWorkbookMinimized(wb, true);
    setWorkbookVisibility(wb, 'hidden');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.bookViews?.[0]?.showHorizontalScroll).toBe(false);
    expect(wb2.bookViews?.[0]?.showVerticalScroll).toBe(false);
    expect(wb2.bookViews?.[0]?.minimized).toBe(true);
    expect(wb2.bookViews?.[0]?.visibility).toBe('hidden');
  });
});
