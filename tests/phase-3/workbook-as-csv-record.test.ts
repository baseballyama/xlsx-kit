// Tests for getWorkbookAsCsvRecord — sheet title → CSV string mapping.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getWorkbookAsCsvRecord,
} from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('getWorkbookAsCsvRecord', () => {
  it('returns one CSV string per worksheet keyed by title', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'a1');
    setCell(a, 1, 2, 'a2');
    setCell(b, 1, 1, 'b1');
    expect(getWorkbookAsCsvRecord(wb)).toEqual({ A: 'a1,a2', B: 'b1' });
  });

  it('includes empty worksheets with ""', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Empty');
    const data = addWorksheet(wb, 'Data');
    setCell(data, 1, 1, 'x');
    expect(getWorkbookAsCsvRecord(wb)).toEqual({ Empty: '', Data: 'x' });
  });

  it('returns {} for a workbook with no worksheets', () => {
    const wb = createWorkbook();
    expect(getWorkbookAsCsvRecord(wb)).toEqual({});
  });

  it('skips chartsheets (they hold no cells)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'x');
    addChartsheet(wb, 'Chart1');
    expect(getWorkbookAsCsvRecord(wb)).toEqual({ Data: 'x' });
  });

  it('forwards opts.delimiter to getWorksheetAsCsv', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    expect(getWorkbookAsCsvRecord(wb, { delimiter: ';' })).toEqual({ Sheet: 'a;b' });
  });
});
