// Tests for parseJsonToRange — JSON array → worksheet range importer.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { parseJsonToRange } from '../../src/worksheet/json';
import { getCell } from '../../src/worksheet/worksheet';

describe('parseJsonToRange', () => {
  it('writes a header row and data rows from a JSON array', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const bbox = parseJsonToRange(
      ws,
      'A1',
      '[{"name":"Alice","age":30},{"name":"Bob","age":25}]',
    );
    expect(bbox).toEqual({ minRow: 1, maxRow: 3, minCol: 1, maxCol: 2 });
    expect(getCell(ws, 1, 1)?.value).toBe('name');
    expect(getCell(ws, 1, 2)?.value).toBe('age');
    expect(getCell(ws, 2, 1)?.value).toBe('Alice');
    expect(getCell(ws, 2, 2)?.value).toBe(30);
    expect(getCell(ws, 3, 1)?.value).toBe('Bob');
    expect(getCell(ws, 3, 2)?.value).toBe(25);
  });

  it('honours opts.keys to override header order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    parseJsonToRange(
      ws,
      'A1',
      [{ name: 'Alice', age: 30 }],
      { keys: ['age', 'name'] },
    );
    expect(getCell(ws, 1, 1)?.value).toBe('age');
    expect(getCell(ws, 1, 2)?.value).toBe('name');
    expect(getCell(ws, 2, 1)?.value).toBe(30);
    expect(getCell(ws, 2, 2)?.value).toBe('Alice');
  });

  it('restores ISO 8601 strings to Date instances', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    parseJsonToRange(ws, 'A1', [{ when: '2026-01-15T03:30:00.000Z' }]);
    const cell = getCell(ws, 2, 1);
    expect(cell?.value).toBeInstanceOf(Date);
    expect((cell?.value as Date).toISOString()).toBe('2026-01-15T03:30:00.000Z');
  });

  it('handles null values and missing keys (both → null)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    parseJsonToRange(
      ws,
      'A1',
      [
        { name: 'Alice', age: 30 },
        { name: 'Bob', age: null },
        { name: 'Carol' },
      ],
    );
    expect(getCell(ws, 2, 2)?.value).toBe(30);
    expect(getCell(ws, 3, 2)?.value ?? null).toBeNull();
    expect(getCell(ws, 4, 2)?.value ?? null).toBeNull();
  });

  it('returns undefined for an empty JSON array (no header written)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(parseJsonToRange(ws, 'A1', '[]')).toBeUndefined();
    expect(parseJsonToRange(ws, 'A1', [])).toBeUndefined();
    expect(getCell(ws, 1, 1)).toBeUndefined();
  });

  it('preserves number / boolean / string mix verbatim', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    parseJsonToRange(
      ws,
      'B2',
      [{ s: 'x', n: 42, b: true, neg: -1.5 }],
    );
    expect(getCell(ws, 2, 2)?.value).toBe('s');
    expect(getCell(ws, 3, 2)?.value).toBe('x');
    expect(getCell(ws, 3, 3)?.value).toBe(42);
    expect(getCell(ws, 3, 4)?.value).toBe(true);
    expect(getCell(ws, 3, 5)?.value).toBe(-1.5);
  });
});
