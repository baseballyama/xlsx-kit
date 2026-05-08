// Tests for getWorkbookStats summary helper.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/xlsx/cell/cell';
import { addDefinedName } from '../../src/xlsx/workbook/defined-names';
import { addChartsheet, addWorksheet, createWorkbook, getWorkbookStats } from '../../src/xlsx/workbook/workbook';
import { setCustomStringProperty } from '../../src/packaging/custom';
import { addUrlHyperlink } from '../../src/xlsx/worksheet/hyperlinks';
import { addExcelTable } from '../../src/xlsx/worksheet/table';
import { mergeCells, setCell, setComment } from '../../src/xlsx/worksheet/worksheet';

describe('getWorkbookStats', () => {
  it('empty workbook → zero everything', () => {
    const wb = createWorkbook();
    expect(getWorkbookStats(wb)).toEqual({
      worksheetCount: 0,
      chartsheetCount: 0,
      cellCount: 0,
      formulaCount: 0,
      commentCount: 0,
      hyperlinkCount: 0,
      mergedRangeCount: 0,
      tableCount: 0,
      definedNameCount: 0,
      customPropertyCount: 0,
    });
  });

  it('counts populated cells + formulas + comments + hyperlinks per worksheet', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 1, 1, 'plain');
    const fc = setCell(a, 1, 2);
    setFormula(fc, 'A1+1');
    setCell(a, 2, 1, 42);
    setComment(a, { ref: 'A1', author: 'Alice', text: 'note' });
    addUrlHyperlink(a, 'A2', 'https://example.com');
    mergeCells(a, 'C1:D1');

    const b = addWorksheet(wb, 'B');
    setCell(b, 1, 1, 'b1');

    const stats = getWorkbookStats(wb);
    expect(stats.worksheetCount).toBe(2);
    expect(stats.cellCount).toBe(4);
    expect(stats.formulaCount).toBe(1);
    expect(stats.commentCount).toBe(1);
    expect(stats.hyperlinkCount).toBe(1);
    expect(stats.mergedRangeCount).toBe(1);
  });

  it('separates chartsheets from worksheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    addChartsheet(wb, 'Chart1');
    const stats = getWorkbookStats(wb);
    expect(stats.worksheetCount).toBe(1);
    expect(stats.chartsheetCount).toBe(1);
  });

  it('counts tables + definedNames + customProperties', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 1, 1, 'h');
    addExcelTable(wb, a, { name: 'Tbl', ref: 'A1:B2', columns: ['c1', 'c2'] });
    addDefinedName(wb, { name: 'Wb', value: '$A$1' });
    addDefinedName(wb, { name: 'Sheet', value: '$A$1', scope: 0 });
    setCustomStringProperty(wb, 'project', 'Apollo');
    const stats = getWorkbookStats(wb);
    expect(stats.tableCount).toBe(1);
    expect(stats.definedNameCount).toBe(2);
    expect(stats.customPropertyCount).toBe(1);
  });
});
