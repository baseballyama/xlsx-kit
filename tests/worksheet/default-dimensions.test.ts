// Tests for setDefaultColumnWidth / setDefaultRowHeight worksheet defaults.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  setDefaultColumnWidth,
  setDefaultRowHeight,
  type Worksheet,
} from '../../src/worksheet/worksheet';

describe('setDefaultColumnWidth', () => {
  it('writes ws.defaultColumnWidth', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setDefaultColumnWidth(ws, 12);
    expect(ws.defaultColumnWidth).toBe(12);
  });

  it('passing undefined clears the field', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setDefaultColumnWidth(ws, 12);
    setDefaultColumnWidth(ws, undefined);
    expect(ws.defaultColumnWidth).toBeUndefined();
  });

  it('rejects negative / NaN', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => setDefaultColumnWidth(ws, -1)).toThrow(/non-negative/);
    expect(() => setDefaultColumnWidth(ws, Number.NaN)).toThrow(/non-negative/);
  });
});

describe('setDefaultRowHeight', () => {
  it('writes ws.defaultRowHeight', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setDefaultRowHeight(ws, 18);
    expect(ws.defaultRowHeight).toBe(18);
  });

  it('passing undefined clears the field', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setDefaultRowHeight(ws, 18);
    setDefaultRowHeight(ws, undefined);
    expect(ws.defaultRowHeight).toBeUndefined();
  });

  it('rejects negative / NaN', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => setDefaultRowHeight(ws, -5)).toThrow(/non-negative/);
    expect(() => setDefaultRowHeight(ws, Number.NaN)).toThrow(/non-negative/);
  });
});

describe('default dimensions round-trip', () => {
  it('both defaults survive saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'D');
    setDefaultColumnWidth(ws, 14);
    setDefaultRowHeight(ws, 22);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb2.sheets[0]?.sheet;
    if (!sheet || !('rows' in sheet)) throw new Error('expected worksheet');
    const ws2 = sheet as Worksheet;
    expect(ws2.defaultColumnWidth).toBe(14);
    expect(ws2.defaultRowHeight).toBe(22);
  });
});
