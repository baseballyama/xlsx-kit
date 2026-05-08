// Tests for createWorkbookFromJsonString — JSON document → fresh Workbook factory.

import { describe, expect, it } from 'vitest';
import {
  createWorkbookFromJsonString,
  getSheet,
  sheetNames,
} from '../../src/workbook/workbook';
import { getCell } from '../../src/worksheet/worksheet';

describe('createWorkbookFromJsonString', () => {
  it('builds a workbook with one sheet per top-level key', () => {
    const wb = createWorkbookFromJsonString(
      '{"A":[{"name":"Alice"}],"B":[{"name":"Bob"}]}',
    );
    expect(sheetNames(wb)).toEqual(['A', 'B']);
    const a = getSheet(wb, 'A');
    const b = getSheet(wb, 'B');
    if (!a || !b) throw new Error('expected both sheets');
    expect(getCell(a, 2, 1)?.value).toBe('Alice');
    expect(getCell(b, 2, 1)?.value).toBe('Bob');
  });

  it('produces a workbook with a single Sheet1 fallback for "{}"', () => {
    const wb = createWorkbookFromJsonString('{}');
    expect(sheetNames(wb)).toEqual(['Sheet1']);
  });

  it('honours opts.fallbackSheetTitle for the empty case', () => {
    const wb = createWorkbookFromJsonString({}, { fallbackSheetTitle: 'Empty' });
    expect(sheetNames(wb)).toEqual(['Empty']);
  });

  it('forwards opts.topLeft to the inner parseJsonStringToWorkbook call', () => {
    const wb = createWorkbookFromJsonString(
      { Data: [{ k: 'v' }] },
      { topLeft: 'B2' },
    );
    const ws = getSheet(wb, 'Data');
    if (!ws) throw new Error('expected Data sheet');
    expect(getCell(ws, 2, 2)?.value).toBe('k');
    expect(getCell(ws, 3, 2)?.value).toBe('v');
  });
});
