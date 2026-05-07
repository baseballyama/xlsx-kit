// Tests for parseCsv / parseCsvToRange — RFC 4180 parser → worksheet write.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { getRangeAsCsv, parseCsv, parseCsvToRange } from '../../src/worksheet/csv';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('parseCsv', () => {
  it('parses a simple comma-delimited CSV', () => {
    expect(parseCsv('a,b,c\n1,2,3')).toEqual([
      ['a', 'b', 'c'],
      ['1', '2', '3'],
    ]);
  });

  it('handles quoted fields with embedded delimiter and quotes (RFC 4180)', () => {
    expect(parseCsv('"a, b","has ""quotes""",plain')).toEqual([['a, b', 'has "quotes"', 'plain']]);
  });

  it('handles embedded newlines inside quoted fields', () => {
    expect(parseCsv('"line1\nline2",plain')).toEqual([['line1\nline2', 'plain']]);
  });

  it('does not emit a trailing empty row when input ends with \\n', () => {
    expect(parseCsv('a,b\n1,2\n')).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ]);
  });

  it('returns [] for empty input', () => {
    expect(parseCsv('')).toEqual([]);
  });
});

describe('parseCsvToRange', () => {
  it('writes the parsed grid starting at the anchor', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const result = parseCsvToRange(ws, 'A1', 'a,b\nc,d');
    expect(result).toEqual({ minRow: 1, maxRow: 2, minCol: 1, maxCol: 2 });
    expect(cellAt(ws, 1, 1).value).toBe('a');
    expect(cellAt(ws, 2, 2).value).toBe('d');
  });

  it('coerceTypes: parses numeric / boolean strings to native types', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    parseCsvToRange(ws, 'A1', 'name,age,active\nAlice,30,true\nBob,25,false', {
      coerceTypes: true,
    });
    expect(cellAt(ws, 2, 1).value).toBe('Alice');
    expect(cellAt(ws, 2, 2).value).toBe(30);
    expect(cellAt(ws, 2, 3).value).toBe(true);
    expect(cellAt(ws, 3, 3).value).toBe(false);
  });

  it('round-trips through getRangeAsCsv (with coerceTypes off)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const original = 'name,age\nAlice,30\nBob,25';
    parseCsvToRange(ws, 'A1', original);
    expect(getRangeAsCsv(ws, 'A1:B3')).toBe(original);
  });

  it('honours opts.delimiter for parse', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    parseCsvToRange(ws, 'A1', 'a;b\n1;2', { delimiter: ';' });
    expect(cellAt(ws, 1, 1).value).toBe('a');
    expect(cellAt(ws, 1, 2).value).toBe('b');
    expect(cellAt(ws, 2, 1).value).toBe('1');
  });

  it('returns undefined for empty input', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(parseCsvToRange(ws, 'A1', '')).toBeUndefined();
    expect(ws.rows.size).toBe(0);
  });
});
