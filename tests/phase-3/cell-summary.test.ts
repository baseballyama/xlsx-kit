// Tests for getCellSummary — debug-friendly per-cell snapshot.

import { describe, expect, it } from 'vitest';
import { setBold } from '../../src/xlsx/styles/cell-style';
import { addWorksheet, createWorkbook, getCellSummary } from '../../src/xlsx/workbook/workbook';
import {
  addConditionalFormatting,
  addDataValidation,
  addTable,
  mergeCells,
  setCell,
  setComment,
  setHyperlink,
} from '../../src/xlsx/worksheet/worksheet';
import { parseMultiCellRange } from '../../src/xlsx/worksheet/cell-range';
import { makeCfRule } from '../../src/xlsx/worksheet/conditional-formatting';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('getCellSummary', () => {
  it('returns a value-only summary for an unstyled cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'plain');
    const s = getCellSummary(wb, 'A', 'A1');
    expect(s.exists).toBe(true);
    expect(s.value).toBe('plain');
    expect(s.styleId).toBe(0);
    expect(s.numberFormat).toBe('General');
    expect(s.font.name).toBe('Calibri');
    expect(s.hyperlink).toBeUndefined();
    expect(s.comment).toBeUndefined();
    expect(s.mergedRange).toBeUndefined();
    expect(s.inTables).toEqual([]);
    expect(s.inDataValidations).toBe(0);
    expect(s.inConditionalFormatting).toBe(0);
  });

  it('reflects styles + hyperlink + comment when populated', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'styled');
    setBold(wb, cellAt(ws, 1, 1));
    setHyperlink(ws, 'A1', { target: 'https://example.com', display: 'click' });
    setComment(ws, { ref: 'A1', text: 'note', author: 'me' });
    const s = getCellSummary(wb, 'A', 'A1');
    expect(s.styleId).not.toBe(0);
    expect(s.font.bold).toBe(true);
    expect(s.hyperlink?.target).toBe('https://example.com');
    expect(s.comment?.text).toBe('note');
  });

  it('reports the merged range for a cell inside a merge', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'tl');
    mergeCells(ws, 'A1:B2');
    const s = getCellSummary(wb, 'A', 'B2');
    expect(s.mergedRange).toBe('A1:B2');
  });

  it('reports table membership and DV / CF counts', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'h');
    setCell(ws, 2, 1, 'd');
    addTable(ws, {
      id: 1,
      displayName: 'T',
      ref: 'A1:A2',
      columns: [{ id: 1, name: 'h' }],
    });
    addDataValidation(ws, {
      type: 'list',
      formula1: '"a,b,c"',
      sqref: parseMultiCellRange('A1:A10'),
    });
    addConditionalFormatting(ws, {
      sqref: parseMultiCellRange('A1:A10'),
      rules: [makeCfRule({ type: 'cellIs', operator: 'equal', formulas: ['"d"'], dxfId: 0, priority: 1 })],
    });
    const s = getCellSummary(wb, 'A', 'A2');
    expect(s.inTables).toEqual(['T']);
    expect(s.inDataValidations).toBe(1);
    expect(s.inConditionalFormatting).toBe(1);
  });

  it('throws when sheetTitle does not resolve', () => {
    const wb = createWorkbook();
    expect(() => getCellSummary(wb, 'Missing', 'A1')).toThrow(/sheet/);
  });

  it('returns exists:false + default styles for an unmaterialised cell coordinate', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 5, 5, 'far'); // populate elsewhere so the sheet exists
    const s = getCellSummary(wb, 'A', 'A1');
    expect(s.exists).toBe(false);
    expect(s.value).toBeUndefined();
    expect(s.styleId).toBe(0);
    expect(s.numberFormat).toBe('General');
  });
});
