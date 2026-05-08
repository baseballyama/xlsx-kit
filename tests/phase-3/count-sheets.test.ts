// Tests for countSheets — sheet count with optional kind/state filters.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  countSheets,
  createWorkbook,
  hideSheet,
} from '../../src/workbook/workbook';

describe('countSheets', () => {
  it('returns the total sheet count by default', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'a');
    addChartsheet(wb, 'b');
    addWorksheet(wb, 'c');
    expect(countSheets(wb)).toBe(3);
  });

  it('opts.kind narrows to worksheets only', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'a');
    addChartsheet(wb, 'b');
    addWorksheet(wb, 'c');
    expect(countSheets(wb, { kind: 'worksheet' })).toBe(2);
  });

  it('opts.kind narrows to chartsheets only', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'a');
    addChartsheet(wb, 'b');
    addChartsheet(wb, 'c');
    expect(countSheets(wb, { kind: 'chartsheet' })).toBe(2);
  });

  it('opts.state narrows by state', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'visible1');
    addWorksheet(wb, 'hidden1');
    addWorksheet(wb, 'visible2');
    hideSheet(wb, 'hidden1');
    expect(countSheets(wb, { state: 'visible' })).toBe(2);
    expect(countSheets(wb, { state: 'hidden' })).toBe(1);
  });

  it('returns 0 for an empty workbook', () => {
    expect(countSheets(createWorkbook())).toBe(0);
  });
});
