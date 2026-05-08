// Tests for iterVisibleWorksheets / iterWorksheetsByState.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  hideSheet,
  iterVisibleWorksheets,
  iterWorksheetsByState,
  setSheetState,
  veryHideSheet,
} from '../../src/workbook/workbook';

describe('iterVisibleWorksheets', () => {
  it('skips hidden / veryHidden worksheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    hideSheet(wb, 'B');
    veryHideSheet(wb, 'C');
    expect([...iterVisibleWorksheets(wb)].map((s) => s.title)).toEqual(['A']);
  });

  it('skips chartsheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addChartsheet(wb, 'Chart');
    expect([...iterVisibleWorksheets(wb)].map((s) => s.title)).toEqual(['A']);
  });

  it('empty workbook → empty', () => {
    const wb = createWorkbook();
    expect([...iterVisibleWorksheets(wb)]).toEqual([]);
  });
});

describe('iterWorksheetsByState', () => {
  it('hidden filter yields only hidden sheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    hideSheet(wb, 'B');
    veryHideSheet(wb, 'C');
    expect([...iterWorksheetsByState(wb, 'hidden')].map((s) => s.title)).toEqual(['B']);
    expect([...iterWorksheetsByState(wb, 'veryHidden')].map((s) => s.title)).toEqual(['C']);
    expect([...iterWorksheetsByState(wb, 'visible')].map((s) => s.title)).toEqual(['A']);
  });

  it('setSheetState updates the iter result', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect([...iterWorksheetsByState(wb, 'hidden')]).toEqual([]);
    setSheetState(wb, 'A', 'hidden');
    expect([...iterWorksheetsByState(wb, 'hidden')].map((s) => s.title)).toEqual(['A']);
  });
});
