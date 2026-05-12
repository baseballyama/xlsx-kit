// Tests for the sheet visibility helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import {
  addWorksheet,
  createWorkbook,
  getSheetState,
  hideSheet,
  setSheetState,
  showSheet,
  veryHideSheet,
} from '../../src/workbook/workbook';

describe('sheet visibility helpers', () => {
  it('newly created sheets default to "visible"', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(getSheetState(wb, 'A')).toBe('visible');
  });

  it('hideSheet / showSheet flip back and forth', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B'); // keep one visible so hiding A doesn't trigger the last-visible guard
    hideSheet(wb, 'A');
    expect(getSheetState(wb, 'A')).toBe('hidden');
    showSheet(wb, 'A');
    expect(getSheetState(wb, 'A')).toBe('visible');
  });

  it('veryHideSheet sets veryHidden state', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B'); // keep one visible
    veryHideSheet(wb, 'A');
    expect(getSheetState(wb, 'A')).toBe('veryHidden');
  });

  it('setSheetState writes the requested state', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B'); // keep one visible
    setSheetState(wb, 'A', 'hidden');
    expect(getSheetState(wb, 'A')).toBe('hidden');
  });

  it('throws on unknown sheet title', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    expect(() => hideSheet(wb, 'Missing')).toThrow();
    expect(() => getSheetState(wb, 'Missing')).toThrow();
    expect(() => setSheetState(wb, 'Missing', 'hidden')).toThrow();
  });

  it('refuses to hide the last visible sheet', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    hideSheet(wb, 'B');
    expect(() => hideSheet(wb, 'A')).toThrow(/last visible sheet/);
    // Excel would refuse to open the workbook if both were hidden; the guard
    // keeps the original state intact.
    expect(getSheetState(wb, 'A')).toBe('visible');
  });

  it('round-trips state through save / load', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    addWorksheet(wb, 'D'); // keep at least one visible
    hideSheet(wb, 'B');
    veryHideSheet(wb, 'C');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(getSheetState(wb2, 'A')).toBe('visible');
    expect(getSheetState(wb2, 'B')).toBe('hidden');
    expect(getSheetState(wb2, 'C')).toBe('veryHidden');
    expect(getSheetState(wb2, 'D')).toBe('visible');
  });
});