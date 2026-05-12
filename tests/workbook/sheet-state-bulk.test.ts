// Tests for setSheetStates / showAllSheets bulk helpers.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  getSheetState,
  hideSheet,
  setSheetStates,
  showAllSheets,
  veryHideSheet,
} from '../../src/workbook/workbook';

describe('setSheetStates', () => {
  it('updates many sheets at once', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    setSheetStates(wb, { A: 'hidden', B: 'veryHidden', C: 'visible' });
    expect(getSheetState(wb, 'A')).toBe('hidden');
    expect(getSheetState(wb, 'B')).toBe('veryHidden');
    expect(getSheetState(wb, 'C')).toBe('visible');
  });

  it('throws when a title is missing', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B'); // keep a visible companion so hiding A doesn't trip the last-visible guard
    expect(() => setSheetStates(wb, { A: 'hidden', Missing: 'hidden' })).toThrow(/no sheet named/);
  });
});

describe('showAllSheets', () => {
  it('flips every hidden / veryHidden sheet to visible and returns the count', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    hideSheet(wb, 'B');
    veryHideSheet(wb, 'C');
    expect(showAllSheets(wb)).toBe(2);
    expect(getSheetState(wb, 'A')).toBe('visible');
    expect(getSheetState(wb, 'B')).toBe('visible');
    expect(getSheetState(wb, 'C')).toBe('visible');
  });

  it('returns 0 when nothing was hidden', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(showAllSheets(wb)).toBe(0);
  });
});
