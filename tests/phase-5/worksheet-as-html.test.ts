// Tests for getWorksheetAsHtml — whole-sheet HTML shortcut.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { getWorksheetAsHtml } from '../../src/worksheet/html';
import { mergeCells, setCell } from '../../src/worksheet/worksheet';

describe('getWorksheetAsHtml', () => {
  it('returns the data extent as an HTML <table>', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    expect(getWorksheetAsHtml(wb, ws)).toBe(
      '<table><tr><td>name</td><td>age</td></tr><tr><td>Alice</td><td>30</td></tr></table>',
    );
  });

  it('returns "" for an empty worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getWorksheetAsHtml(wb, ws)).toBe('');
  });

  it('uses the data extent (sparse layout)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'tl');
    setCell(ws, 3, 3, 'br');
    const html = getWorksheetAsHtml(wb, ws);
    // 3 rows × 3 cols
    const rowCount = (html.match(/<tr>/g) ?? []).length;
    expect(rowCount).toBe(3);
  });

  it('includes merge collapse via worksheetToHtml', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'h');
    mergeCells(ws, 'A1:B2');
    setCell(ws, 3, 1, 'a'); setCell(ws, 3, 2, 'b');
    expect(getWorksheetAsHtml(wb, ws)).toContain('rowspan="2"');
  });
});
