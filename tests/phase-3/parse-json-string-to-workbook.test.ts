// Tests for parseJsonStringToWorkbook — JSON document → workbook importer.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  getSheet,
  parseJsonStringToWorkbook,
  sheetNames,
} from '../../src/workbook/workbook';
import { getCell } from '../../src/worksheet/worksheet';

describe('parseJsonStringToWorkbook', () => {
  it('adds one worksheet per top-level key in document order', () => {
    const wb = createWorkbook();
    const created = parseJsonStringToWorkbook(
      wb,
      '{"A":[{"name":"Alice"}],"B":[{"name":"Bob"}]}',
    );
    expect(created).toEqual(['A', 'B']);
    expect(sheetNames(wb)).toEqual(['A', 'B']);
    const a = getSheet(wb, 'A');
    const b = getSheet(wb, 'B');
    if (!a || !b) throw new Error('expected both sheets');
    expect(getCell(a, 2, 1)?.value).toBe('Alice');
    expect(getCell(b, 2, 1)?.value).toBe('Bob');
  });

  it('returns [] and adds nothing for an empty JSON object', () => {
    const wb = createWorkbook();
    expect(parseJsonStringToWorkbook(wb, '{}')).toEqual([]);
    expect(sheetNames(wb)).toEqual([]);
  });

  it('honours opts.topLeft to anchor each sheet', () => {
    const wb = createWorkbook();
    parseJsonStringToWorkbook(
      wb,
      { Data: [{ k: 'v' }] },
      { topLeft: 'C5' },
    );
    const ws = getSheet(wb, 'Data');
    if (!ws) throw new Error('expected Data sheet');
    expect(getCell(ws, 5, 3)?.value).toBe('k');
    expect(getCell(ws, 6, 3)?.value).toBe('v');
  });

  it('honours opts.keys per-sheet for header order override', () => {
    const wb = createWorkbook();
    parseJsonStringToWorkbook(
      wb,
      { People: [{ name: 'Alice', age: 30 }] },
      { keys: { People: ['age', 'name'] } },
    );
    const ws = getSheet(wb, 'People');
    if (!ws) throw new Error('expected People sheet');
    expect(getCell(ws, 1, 1)?.value).toBe('age');
    expect(getCell(ws, 1, 2)?.value).toBe('name');
    expect(getCell(ws, 2, 1)?.value).toBe(30);
    expect(getCell(ws, 2, 2)?.value).toBe('Alice');
  });

  it('throws on title collision; opts.replace overwrites', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    expect(() =>
      parseJsonStringToWorkbook(wb, '{"A":[{"k":"v"}]}'),
    ).toThrow(/already exists/);
    parseJsonStringToWorkbook(wb, '{"A":[{"k":"v"}]}', { replace: true });
    const ws = getSheet(wb, 'A');
    if (!ws) throw new Error('expected sheet A');
    expect(getCell(ws, 2, 1)?.value).toBe('v');
  });
});
