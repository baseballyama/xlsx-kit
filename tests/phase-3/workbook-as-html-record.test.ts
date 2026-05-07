// Tests for getWorkbookAsHtmlRecord — sheet title → HTML mapping.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getWorkbookAsHtmlRecord,
} from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('getWorkbookAsHtmlRecord', () => {
  it('returns one HTML <table> per worksheet keyed by title', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'a-cell');
    setCell(b, 1, 1, 'b-cell');
    const r = getWorkbookAsHtmlRecord(wb);
    expect(Object.keys(r).sort()).toEqual(['A', 'B']);
    expect(r['A']).toContain('<table>');
    expect(r['A']).toContain('a-cell');
    expect(r['B']).toContain('b-cell');
  });

  it('includes empty worksheets as ""', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Empty');
    const data = addWorksheet(wb, 'Data');
    setCell(data, 1, 1, 'x');
    expect(getWorkbookAsHtmlRecord(wb)).toEqual({
      Empty: '',
      Data: '<table><tr><td>x</td></tr></table>',
    });
  });

  it('returns {} for a workbook with no worksheets', () => {
    expect(getWorkbookAsHtmlRecord(createWorkbook())).toEqual({});
  });

  it('skips chartsheets (no cells)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'x');
    addChartsheet(wb, 'Chart1');
    expect(Object.keys(getWorkbookAsHtmlRecord(wb))).toEqual(['Data']);
  });
});
