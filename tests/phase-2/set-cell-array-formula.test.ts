// Tests for setCellArrayFormula.

import { describe, expect, it } from 'vitest';
import { isFormulaValue } from '../../src/cell/cell';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCellArrayFormula } from '../../src/worksheet/worksheet';

describe('setCellArrayFormula', () => {
  it('writes an array formula spanning the supplied ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCellArrayFormula(ws, 1, 1, 'A1:A3', '=TRANSPOSE(B1:D1)');
    expect(isFormulaValue(c.value)).toBe(true);
    if (isFormulaValue(c.value)) {
      expect(c.value.t).toBe('array');
      expect(c.value.formula).toBe('TRANSPOSE(B1:D1)');
      expect(c.value.ref).toBe('A1:A3');
    }
  });

  it('accepts the formula without =', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCellArrayFormula(ws, 1, 1, 'A1:B2', 'SEQUENCE(2, 2)');
    if (isFormulaValue(c.value)) {
      expect(c.value.formula).toBe('SEQUENCE(2, 2)');
    }
  });

  it('passes through cachedValue + styleId', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCellArrayFormula(ws, 1, 1, 'A1:A2', '=ROW(A1:A2)', { cachedValue: 1, styleId: 4 });
    if (isFormulaValue(c.value)) expect(c.value.cachedValue).toBe(1);
    expect(c.styleId).toBe(4);
  });
});
