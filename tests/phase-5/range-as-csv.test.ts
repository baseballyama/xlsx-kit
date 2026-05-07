// Tests for getRangeAsCsv — render a range as CSV.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { getRangeAsCsv } from '../../src/worksheet/csv';
import { setCell } from '../../src/worksheet/worksheet';

describe('getRangeAsCsv', () => {
  it('renders a simple grid as comma-separated rows joined with \\n', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice');
    setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Bob');
    setCell(ws, 3, 2, 25);
    expect(getRangeAsCsv(ws, 'A1:B3')).toBe('name,age\nAlice,30\nBob,25');
  });

  it('quotes fields containing the delimiter, embeds doubled quotes for `"`', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a, b');
    setCell(ws, 1, 2, 'has "quotes"');
    setCell(ws, 1, 3, 'plain');
    expect(getRangeAsCsv(ws, 'A1:C1')).toBe('"a, b","has ""quotes""",plain');
  });

  it('quotes fields containing newlines', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'line1\nline2');
    setCell(ws, 1, 2, 'plain');
    expect(getRangeAsCsv(ws, 'A1:B1')).toBe('"line1\nline2",plain');
  });

  it('renders empty cells as empty fields (preserves column count)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    // col 2 intentionally empty
    setCell(ws, 1, 3, 'c');
    expect(getRangeAsCsv(ws, 'A1:C1')).toBe('a,,c');
  });

  it('honours opts.delimiter (semicolon) and opts.trailingNewline', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'x');
    setCell(ws, 1, 2, 'y');
    expect(getRangeAsCsv(ws, 'A1:B1', { delimiter: ';' })).toBe('x;y');
    expect(getRangeAsCsv(ws, 'A1:B1', { trailingNewline: true })).toBe('x,y\n');
  });

  it('coerces Date / boolean / number cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, new Date('2025-01-01T00:00:00Z'));
    setCell(ws, 1, 2, true);
    setCell(ws, 1, 3, 1.5);
    expect(getRangeAsCsv(ws, 'A1:C1')).toBe('2025-01-01T00:00:00.000Z,true,1.5');
  });
});
