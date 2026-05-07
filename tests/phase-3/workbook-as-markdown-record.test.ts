// Tests for getWorkbookAsMarkdownRecord — sheet title → MD table mapping.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getWorkbookAsMarkdownRecord,
} from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('getWorkbookAsMarkdownRecord', () => {
  it('returns one MD table per worksheet keyed by title', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'a-cell');
    setCell(b, 1, 1, 'b-cell');
    const r = getWorkbookAsMarkdownRecord(wb);
    expect(Object.keys(r).sort()).toEqual(['A', 'B']);
    expect(r['A']).toContain('| a-cell |');
    expect(r['B']).toContain('| b-cell |');
  });

  it('includes empty worksheets as ""', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Empty');
    const data = addWorksheet(wb, 'Data');
    setCell(data, 1, 1, 'x');
    const r = getWorkbookAsMarkdownRecord(wb);
    expect(r['Empty']).toBe('');
    expect(r['Data']).toBe(['| x |', '| --- |'].join('\n'));
  });

  it('returns {} for a workbook with no worksheets', () => {
    expect(getWorkbookAsMarkdownRecord(createWorkbook())).toEqual({});
  });

  it('skips chartsheets (no cells)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'x');
    addChartsheet(wb, 'Chart1');
    expect(Object.keys(getWorkbookAsMarkdownRecord(wb))).toEqual(['Data']);
  });
});
