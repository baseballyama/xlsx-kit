// Tests for worksheetToJson — header-driven JSON renderer.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { worksheetToJson } from '../../src/worksheet/json';
import { setCell } from '../../src/worksheet/worksheet';

describe('worksheetToJson', () => {
  it('renders rows as JSON objects keyed by the header row (single-line by default)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Bob'); setCell(ws, 3, 2, 25);
    const json = worksheetToJson(ws, 'A1:B3');
    expect(json).toBe('[{"name":"Alice","age":30},{"name":"Bob","age":25}]');
    expect(JSON.parse(json)).toEqual([
      { name: 'Alice', age: 30 },
      { name: 'Bob', age: 25 },
    ]);
  });

  it('honours opts.pretty by emitting 2-space indentation with newlines', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'v');
    const pretty = worksheetToJson(ws, 'A1:A2', { pretty: true });
    expect(pretty).toContain('\n');
    expect(pretty).toContain('  "k"');
    expect(JSON.parse(pretty)).toEqual([{ k: 'v' }]);
  });

  it('serialises Date cells as ISO 8601 strings', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'when');
    setCell(ws, 2, 1, new Date('2026-01-15T03:30:00.000Z'));
    expect(worksheetToJson(ws, 'A1:A2')).toBe('[{"when":"2026-01-15T03:30:00.000Z"}]');
  });

  it('uses formula cachedValue when set, falling back to the formula text', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'cached'); setCell(ws, 1, 2, 'uncached');
    const a = setCell(ws, 2, 1);
    setFormula(a, 'A1+1', { cachedValue: 42 });
    const b = setCell(ws, 2, 2);
    setFormula(b, 'B1+1');
    expect(worksheetToJson(ws, 'A1:B2')).toBe('[{"cached":42,"uncached":"B1+1"}]');
  });

  it('concatenates rich-text run text and maps duration / error variants', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'rt'); setCell(ws, 1, 2, 'dur'); setCell(ws, 1, 3, 'err');
    setCell(ws, 2, 1, {
      kind: 'rich-text',
      runs: [{ text: 'Hello ' }, { text: 'World' }],
    });
    setCell(ws, 2, 2, { kind: 'duration', ms: 1500 });
    setCell(ws, 2, 3, { kind: 'error', code: '#REF!' });
    expect(worksheetToJson(ws, 'A1:C2')).toBe(
      '[{"rt":"Hello World","dur":1500,"err":"#REF!"}]',
    );
  });
});
