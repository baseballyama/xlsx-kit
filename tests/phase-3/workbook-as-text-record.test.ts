// Tests for getWorkbookAsTextTableRecord — sheet title → text table mapping.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getWorkbookAsTextTableRecord,
} from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('getWorkbookAsTextTableRecord', () => {
  it('returns one ASCII-art table per worksheet keyed by title', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'a-cell');
    setCell(b, 1, 1, 'b-cell');
    const r = getWorkbookAsTextTableRecord(wb);
    expect(Object.keys(r).sort()).toEqual(['A', 'B']);
    expect(r['A']).toContain('a-cell');
    expect(r['B']).toContain('b-cell');
  });

  it('includes empty worksheets as ""', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Empty');
    const data = addWorksheet(wb, 'Data');
    setCell(data, 1, 1, 'x');
    const r = getWorkbookAsTextTableRecord(wb);
    expect(r['Empty']).toBe('');
    expect(r['Data']).toBe(['| x |', '+---+'].join('\n'));
  });

  it('skips chartsheets (no cells)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'x');
    addChartsheet(wb, 'Chart1');
    expect(Object.keys(getWorkbookAsTextTableRecord(wb))).toEqual(['Data']);
  });
});
