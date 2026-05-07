// Tests for appendRows — bulk version of appendRow.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { appendRow, appendRows } from '../../src/worksheet/worksheet';

describe('appendRows', () => {
  it('appends a 2D array of values and returns {firstRow, lastRow}', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const result = appendRows(ws, [
      ['a', 1],
      ['b', 2],
      ['c', 3],
    ]);
    expect(result).toEqual({ firstRow: 1, lastRow: 3 });
    expect(ws.rows.get(1)?.get(1)?.value).toBe('a');
    expect(ws.rows.get(2)?.get(2)?.value).toBe(2);
    expect(ws.rows.get(3)?.get(1)?.value).toBe('c');
  });

  it('advances the cursor for empty rows so subsequent appends do not overlap', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    appendRows(ws, [['x'], [], ['y']]);
    expect(ws.rows.get(1)?.get(1)?.value).toBe('x');
    expect(ws.rows.get(3)?.get(1)?.value).toBe('y');
    // Row 2 is intentionally absent (empty input row → cursor advanced, no cells)
    expect(ws.rows.get(2)).toBeUndefined();
  });

  it('skips undefined / null entries within a row (no cell created)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    appendRows(ws, [
      ['a', undefined, 'c'],
      [null, 'b', null],
    ]);
    expect(ws.rows.get(1)?.has(2)).toBe(false);
    expect(ws.rows.get(1)?.get(3)?.value).toBe('c');
    expect(ws.rows.get(2)?.has(1)).toBe(false);
    expect(ws.rows.get(2)?.get(2)?.value).toBe('b');
  });

  it('respects the existing append cursor — appends after the last appended row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    appendRow(ws, ['existing']); // → row 1
    const result = appendRows(ws, [['next'], ['after']]);
    expect(result).toEqual({ firstRow: 2, lastRow: 3 });
    expect(ws.rows.get(2)?.get(1)?.value).toBe('next');
    expect(ws.rows.get(3)?.get(1)?.value).toBe('after');
  });

  it('returns a marker no-op result on empty input (lastRow = firstRow - 1)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const result = appendRows(ws, []);
    expect(result.firstRow).toBe(1);
    expect(result.lastRow).toBe(0);
    // No cells were materialised by the call.
    expect(ws.rows.size).toBe(0);
  });
});
