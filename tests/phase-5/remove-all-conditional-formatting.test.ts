// Tests for removeAllConditionalFormatting.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { addCellIsRule, addColorScaleRule } from '../../src/worksheet/conditional-formatting';
import {
  getConditionalFormatting,
  removeAllConditionalFormatting,
} from '../../src/worksheet/worksheet';

describe('removeAllConditionalFormatting', () => {
  it('drops every CF block and returns the count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    addCellIsRule(ws, 'A1:A10', { operator: 'greaterThan', formula1: '0' });
    addColorScaleRule(ws, 'B1:B5', {
      cfvos: [{ type: 'min' }, { type: 'max' }],
      colors: ['FFFF0000', 'FF00FF00'],
    });
    expect(removeAllConditionalFormatting(ws)).toBe(2);
    expect(getConditionalFormatting(ws)).toEqual([]);
  });

  it('returns 0 when no CF blocks exist', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(removeAllConditionalFormatting(ws)).toBe(0);
  });
});
