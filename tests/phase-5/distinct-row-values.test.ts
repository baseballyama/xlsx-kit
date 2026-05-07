// Tests for getDistinctValuesInRow — row variant of getDistinctValuesInColumn.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  getDistinctValuesInRow,
  setCell,
} from '../../src/worksheet/worksheet';

describe('getDistinctValuesInRow', () => {
  it('dedupes values and preserves first-seen column order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'apple');
    setCell(ws, 1, 2, 'banana');
    setCell(ws, 1, 3, 'apple'); // dup
    setCell(ws, 1, 4, 'cherry');
    setCell(ws, 1, 5, 'banana'); // dup
    expect(getDistinctValuesInRow(ws, 1)).toEqual(['apple', 'banana', 'cherry']);
  });

  it('handles mixed types as distinct (number 1 ≠ string "1")', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setCell(ws, 1, 2, '1');
    setCell(ws, 1, 3, true);
    setCell(ws, 1, 4, 1);
    expect(getDistinctValuesInRow(ws, 1)).toEqual([1, '1', true]);
  });

  it('returns [] for an empty / absent row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 5, 1, 'x'); // populate row 5; row 1 still empty
    expect(getDistinctValuesInRow(ws, 1)).toEqual([]);
  });

  it('skipNull drops null values', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, null);
    setCell(ws, 1, 3, 'b');
    expect(getDistinctValuesInRow(ws, 1)).toEqual(['a', null, 'b']);
    expect(getDistinctValuesInRow(ws, 1, { skipNull: true })).toEqual(['a', 'b']);
  });

  it('skipFormulas drops FormulaValue cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'plain');
    const c = setCell(ws, 1, 2);
    setFormula(c, 'A1+1');
    setCell(ws, 1, 3, 'tail');
    expect(getDistinctValuesInRow(ws, 1, { skipFormulas: true })).toEqual(['plain', 'tail']);
  });
});
