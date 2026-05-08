// Tests for getWorkbookAsJsonString — combined whole-workbook JSON document.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getWorkbookAsJsonString,
} from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('getWorkbookAsJsonString', () => {
  it('returns one combined JSON document keyed by sheet title', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'name'); setCell(a, 2, 1, 'Alice');
    setCell(b, 1, 1, 'name'); setCell(b, 2, 1, 'Bob');
    expect(getWorkbookAsJsonString(wb)).toBe(
      '{"A":[{"name":"Alice"}],"B":[{"name":"Bob"}]}',
    );
  });

  it('returns "{}" for a workbook with no worksheets', () => {
    expect(getWorkbookAsJsonString(createWorkbook())).toBe('{}');
  });

  it('represents empty worksheets as an empty array', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Empty');
    const data = addWorksheet(wb, 'Data');
    setCell(data, 1, 1, 'k');
    setCell(data, 2, 1, 'v');
    const parsed = JSON.parse(getWorkbookAsJsonString(wb));
    expect(parsed).toEqual({ Empty: [], Data: [{ k: 'v' }] });
  });

  it('honours opts.pretty (2-space indented document)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'v');
    const pretty = getWorkbookAsJsonString(wb, { pretty: true });
    expect(pretty).toContain('\n');
    expect(pretty).toContain('  "A"');
    expect(JSON.parse(pretty)).toEqual({ A: [{ k: 'v' }] });
  });

  it('skips chartsheets and applies the same cell coercion as worksheetToJson', () => {
    const wb = createWorkbook();
    const data = addWorksheet(wb, 'Data');
    setCell(data, 1, 1, 'date'); setCell(data, 1, 2, 'formula'); setCell(data, 1, 3, 'rt');
    setCell(data, 2, 1, new Date('2026-01-15T03:30:00.000Z'));
    const fc = setCell(data, 2, 2);
    setFormula(fc, 'A1+1', { cachedValue: 42 });
    setCell(data, 2, 3, {
      kind: 'rich-text',
      runs: [{ text: 'Hi ' }, { text: 'there' }],
    });
    addChartsheet(wb, 'Chart1');
    const parsed = JSON.parse(getWorkbookAsJsonString(wb));
    expect(Object.keys(parsed)).toEqual(['Data']);
    expect(parsed['Data']).toEqual([
      { date: '2026-01-15T03:30:00.000Z', formula: 42, rt: 'Hi there' },
    ]);
  });
});
