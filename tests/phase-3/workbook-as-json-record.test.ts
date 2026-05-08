// Tests for getWorkbookAsJsonRecord — sheet title → JSON string mapping.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getWorkbookAsJsonRecord,
} from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('getWorkbookAsJsonRecord', () => {
  it('returns one JSON string per worksheet keyed by title', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'name');
    setCell(a, 2, 1, 'Alice');
    setCell(b, 1, 1, 'name');
    setCell(b, 2, 1, 'Bob');
    const r = getWorkbookAsJsonRecord(wb);
    expect(Object.keys(r).sort()).toEqual(['A', 'B']);
    expect(JSON.parse(r['A'] ?? '')).toEqual([{ name: 'Alice' }]);
    expect(JSON.parse(r['B'] ?? '')).toEqual([{ name: 'Bob' }]);
  });

  it('includes empty worksheets as "[]"', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Empty');
    const data = addWorksheet(wb, 'Data');
    setCell(data, 1, 1, 'k');
    setCell(data, 2, 1, 'v');
    const r = getWorkbookAsJsonRecord(wb);
    expect(r['Empty']).toBe('[]');
    expect(r['Data']).toBe('[{"k":"v"}]');
  });

  it('returns {} for a workbook with no worksheets', () => {
    expect(getWorkbookAsJsonRecord(createWorkbook())).toEqual({});
  });

  it('skips chartsheets (no cells)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'v');
    addChartsheet(wb, 'Chart1');
    expect(Object.keys(getWorkbookAsJsonRecord(wb))).toEqual(['Data']);
  });
});
