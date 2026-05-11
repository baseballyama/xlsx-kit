// Tests for setCellFormula.

import { describe, expect, it } from 'vitest';
import { getCachedFormulaValue, getFormulaText, isFormulaValue } from '../../src/cell/cell';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCellFormula } from '../../src/worksheet/worksheet';

describe('setCellFormula', () => {
  it('writes a normal formula and strips the leading =', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCellFormula(ws, 1, 1, '=A2+B2');
    expect(isFormulaValue(c.value)).toBe(true);
    expect(getFormulaText(c)).toBe('A2+B2');
  });

  it('accepts the formula without =', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCellFormula(ws, 1, 1, 'SUM(A:A)');
    expect(getFormulaText(c)).toBe('SUM(A:A)');
  });

  it('passes through cachedValue', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCellFormula(ws, 1, 1, '=A2+1', { cachedValue: 42 });
    expect(getCachedFormulaValue(c)).toBe(42);
  });

  it('honours styleId', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCellFormula(ws, 1, 1, '=1', { styleId: 3 });
    expect(c.styleId).toBe(3);
  });
});
