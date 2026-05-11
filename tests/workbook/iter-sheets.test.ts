// Tests for iterWorksheets / iterChartsheets / listWorksheets / listChartsheets.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  iterChartsheets,
  iterWorksheets,
  listChartsheets,
  listWorksheets,
} from '../../src/workbook/workbook';

describe('iterWorksheets / listWorksheets', () => {
  it('iterates every worksheet in tab-strip order', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addChartsheet(wb, 'Chart1');
    addWorksheet(wb, 'B');
    expect([...iterWorksheets(wb)].map((s) => s.title)).toEqual(['A', 'B']);
    expect(listWorksheets(wb).map((s) => s.title)).toEqual(['A', 'B']);
  });

  it('empty workbook → empty iteration', () => {
    const wb = createWorkbook();
    expect(listWorksheets(wb)).toEqual([]);
  });
});

describe('iterChartsheets / listChartsheets', () => {
  it('skips regular worksheets and yields only chartsheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addChartsheet(wb, 'Chart1');
    addChartsheet(wb, 'Chart2');
    addWorksheet(wb, 'B');
    expect([...iterChartsheets(wb)].map((s) => s.title)).toEqual(['Chart1', 'Chart2']);
    expect(listChartsheets(wb).map((s) => s.title)).toEqual(['Chart1', 'Chart2']);
  });

  it('workbook with no chartsheets → empty', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(listChartsheets(wb)).toEqual([]);
  });
});
