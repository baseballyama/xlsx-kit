// Tests for getHeaders — header row → string[].

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { getHeaders, setCell } from '../../src/worksheet/worksheet';

describe('getHeaders', () => {
  it('returns the header row in column order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name');
    setCell(ws, 1, 2, 'age');
    setCell(ws, 1, 3, 'role');
    expect(getHeaders(ws, 'A1:C1')).toEqual(['name', 'age', 'role']);
  });

  it('coerces non-string header cells to strings', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 42);
    setCell(ws, 1, 2, true);
    expect(getHeaders(ws, 'A1:B1')).toEqual(['42', 'true']);
  });

  it('renders empty / unmaterialised header cells as ""', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    // col 2 left empty
    setCell(ws, 1, 3, 'c');
    expect(getHeaders(ws, 'A1:C1')).toEqual(['a', '', 'c']);
  });

  it('returns an array of length maxCol - minCol + 1', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getHeaders(ws, 'A1:E1').length).toBe(5);
  });
});
