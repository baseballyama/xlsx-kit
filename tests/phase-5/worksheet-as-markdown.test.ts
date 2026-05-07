// Tests for getWorksheetAsMarkdownTable — whole-sheet markdown shortcut.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { getWorksheetAsMarkdownTable } from '../../src/worksheet/markdown';
import { mergeCells, setCell } from '../../src/worksheet/worksheet';

describe('getWorksheetAsMarkdownTable', () => {
  it('returns the data extent as a GFM table', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    expect(getWorksheetAsMarkdownTable(ws)).toBe(
      ['| name | age |', '| --- | --- |', '| Alice | 30 |'].join('\n'),
    );
  });

  it('returns "" for an empty worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getWorksheetAsMarkdownTable(ws)).toBe('');
  });

  it('uses the data extent (sparse layout)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'tl');
    setCell(ws, 3, 3, 'br');
    const md = getWorksheetAsMarkdownTable(ws);
    expect(md.split('\n').length).toBe(4); // header + sep + 2 data rows
  });

  it('flattens merged ranges', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'h'); setCell(ws, 1, 2, 'k');
    mergeCells(ws, 'A2:B2');
    setCell(ws, 2, 1, 'merged');
    expect(getWorksheetAsMarkdownTable(ws)).toContain('| merged |  |');
  });
});
