// End-to-end round-trip: getWorkbookAsJsonString → parseJsonStringToWorkbook.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  getSheet,
  getWorkbookAsJsonString,
  parseJsonStringToWorkbook,
  sheetNames,
} from '../../src/workbook/workbook';
import { getCell, setCell } from '../../src/worksheet/worksheet';

describe('getWorkbookAsJsonString + parseJsonStringToWorkbook round-trip', () => {
  it('round-trips multi-sheet cell values verbatim', () => {
    const src = createWorkbook();
    const a = addWorksheet(src, 'A');
    setCell(a, 1, 1, 'name'); setCell(a, 1, 2, 'age');
    setCell(a, 2, 1, 'Alice'); setCell(a, 2, 2, 30);
    const b = addWorksheet(src, 'B');
    setCell(b, 1, 1, 'when');
    setCell(b, 2, 1, new Date('2026-01-15T03:30:00.000Z'));

    const json = getWorkbookAsJsonString(src);
    const dst = createWorkbook();
    parseJsonStringToWorkbook(dst, json);

    expect(sheetNames(dst)).toEqual(['A', 'B']);
    const dstA = getSheet(dst, 'A');
    const dstB = getSheet(dst, 'B');
    if (!dstA || !dstB) throw new Error('expected both restored sheets');
    expect(getCell(dstA, 2, 1)?.value).toBe('Alice');
    expect(getCell(dstA, 2, 2)?.value).toBe(30);
    const restoredDate = getCell(dstB, 2, 1)?.value;
    expect(restoredDate).toBeInstanceOf(Date);
    expect((restoredDate as Date).toISOString()).toBe('2026-01-15T03:30:00.000Z');
  });
});
