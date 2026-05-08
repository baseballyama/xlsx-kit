// Tests for getAllConditionalFormatting.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getAllConditionalFormatting,
} from '../../src/xlsx/workbook/workbook';
import { addCellIsRule, addColorScaleRule } from '../../src/xlsx/worksheet/conditional-formatting';

describe('getAllConditionalFormatting', () => {
  it('aggregates CF blocks across every worksheet in tab-strip order', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    addCellIsRule(a, 'A1:A10', { operator: 'greaterThan', formula1: '0' });
    addColorScaleRule(b, 'B1:B5', {
      cfvos: [{ type: 'min' }, { type: 'max' }],
      colors: ['FFFF0000', 'FF00FF00'],
    });
    addCellIsRule(a, 'C1:C5', { operator: 'lessThan', formula1: '100' });
    const out = getAllConditionalFormatting(wb).map(
      ({ sheet, formatting }) => `${sheet.title}:${formatting.rules[0]?.type}`,
    );
    expect(out).toEqual(['A:cellIs', 'A:cellIs', 'B:colorScale']);
  });

  it('skips chartsheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A'); // empty
    addChartsheet(wb, 'Chart');
    expect(getAllConditionalFormatting(wb)).toEqual([]);
  });

  it('empty workbook → empty array', () => {
    const wb = createWorkbook();
    expect(getAllConditionalFormatting(wb)).toEqual([]);
  });
});
