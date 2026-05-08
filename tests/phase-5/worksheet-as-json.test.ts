// Tests for getWorksheetAsJson — whole-sheet JSON shortcut.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { getWorksheetAsJson } from '../../src/worksheet/json';
import { setCell } from '../../src/worksheet/worksheet';

describe('getWorksheetAsJson', () => {
  it('returns the data extent serialised as JSON', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Bob'); setCell(ws, 3, 2, 25);
    expect(getWorksheetAsJson(ws)).toBe(
      '[{"name":"Alice","age":30},{"name":"Bob","age":25}]',
    );
  });

  it('returns "[]" for an empty worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getWorksheetAsJson(ws)).toBe('[]');
  });

  it('uses the data extent on a sparse layout', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k1'); setCell(ws, 1, 3, 'k2');
    setCell(ws, 3, 3, 'v');
    const parsed = JSON.parse(getWorksheetAsJson(ws)) as Array<Record<string, unknown>>;
    expect(parsed).toHaveLength(2);
    const last = parsed[1];
    if (!last) throw new Error('expected at least 2 rows');
    expect(last['k2']).toBe('v');
    expect(last['k1']).toBeNull();
  });

  it('forwards opts.pretty to worksheetToJson', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'v');
    const pretty = getWorksheetAsJson(ws, { pretty: true });
    expect(pretty).toContain('\n');
    expect(pretty).toContain('  "k"');
    expect(JSON.parse(pretty)).toEqual([{ k: 'v' }]);
  });
});
