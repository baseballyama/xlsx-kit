// Tests for getSheetTitles — sheet title array with optional kind/state filters.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getSheetTitles,
  hideSheet,
} from '../../src/xlsx/workbook/workbook';

describe('getSheetTitles', () => {
  it('returns every sheet title in tab-strip order by default', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'first');
    addChartsheet(wb, 'middle');
    addWorksheet(wb, 'last');
    expect(getSheetTitles(wb)).toEqual(['first', 'middle', 'last']);
  });

  it('opts.kind narrows to worksheets only', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'data');
    addChartsheet(wb, 'chart');
    addWorksheet(wb, 'more');
    expect(getSheetTitles(wb, { kind: 'worksheet' })).toEqual(['data', 'more']);
  });

  it('opts.kind narrows to chartsheets only', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'data');
    addChartsheet(wb, 'chart');
    addChartsheet(wb, 'plot');
    expect(getSheetTitles(wb, { kind: 'chartsheet' })).toEqual(['chart', 'plot']);
  });

  it('opts.state narrows by state', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'visible1');
    addWorksheet(wb, 'hidden1');
    addWorksheet(wb, 'visible2');
    hideSheet(wb, 'hidden1');
    expect(getSheetTitles(wb, { state: 'visible' })).toEqual(['visible1', 'visible2']);
    expect(getSheetTitles(wb, { state: 'hidden' })).toEqual(['hidden1']);
  });

  it('returns [] for an empty workbook', () => {
    expect(getSheetTitles(createWorkbook())).toEqual([]);
  });
});
