// Tests for getDistinctValuesInColumn — column dedup with formula/null filters.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  getDistinctValuesInColumn,
  setCell,
} from '../../src/worksheet/worksheet';

describe('getDistinctValuesInColumn', () => {
  it('dedupes primitive values and preserves first-seen row order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'apple');
    setCell(ws, 2, 1, 'banana');
    setCell(ws, 3, 1, 'apple'); // dup
    setCell(ws, 4, 1, 'cherry');
    setCell(ws, 5, 1, 'banana'); // dup
    expect(getDistinctValuesInColumn(ws, 1)).toEqual(['apple', 'banana', 'cherry']);
  });

  it('handles mixed numeric / boolean / string types as distinct', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setCell(ws, 2, 1, '1'); // string '1' is distinct from number 1
    setCell(ws, 3, 1, true);
    setCell(ws, 4, 1, 1); // dup of number 1
    expect(getDistinctValuesInColumn(ws, 1)).toEqual([1, '1', true]);
  });

  it('returns [] for an empty / absent column', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 5, 'x'); // populate col 5; col 1 still empty
    expect(getDistinctValuesInColumn(ws, 1)).toEqual([]);
  });

  it('skipNull drops null values', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 2, 1, null);
    setCell(ws, 3, 1, 'b');
    expect(getDistinctValuesInColumn(ws, 1)).toEqual(['a', null, 'b']);
    expect(getDistinctValuesInColumn(ws, 1, { skipNull: true })).toEqual(['a', 'b']);
  });

  it('skipFormulas drops FormulaValue cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'plain');
    const c = setCell(ws, 2, 1);
    setFormula(c, 'A1+1');
    setCell(ws, 3, 1, 'tail');
    expect(getDistinctValuesInColumn(ws, 1, { skipFormulas: true })).toEqual(['plain', 'tail']);
  });
});
