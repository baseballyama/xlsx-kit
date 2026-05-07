// Tests for worksheetToMarkdownTable — GFM table renderer.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { worksheetToMarkdownTable } from '../../src/worksheet/markdown';
import { mergeCells, setCell } from '../../src/worksheet/worksheet';

describe('worksheetToMarkdownTable', () => {
  it('renders a header + data rows separated by --- markers', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Bob'); setCell(ws, 3, 2, 25);
    expect(worksheetToMarkdownTable(ws, 'A1:B3')).toBe(
      [
        '| name | age |',
        '| --- | --- |',
        '| Alice | 30 |',
        '| Bob | 25 |',
      ].join('\n'),
    );
  });

  it('escapes pipe characters and newlines inside cell values', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'col');
    setCell(ws, 2, 1, 'a | b');
    setCell(ws, 3, 1, 'line1\nline2');
    const md = worksheetToMarkdownTable(ws, 'A1:A3');
    expect(md).toContain('a \\| b');
    expect(md).toContain('line1<br>line2');
  });

  it('renders empty / unmaterialised cells as empty markdown cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'x');
    // (2, 2) intentionally empty
    const md = worksheetToMarkdownTable(ws, 'A1:B2');
    expect(md.split('\n')[2]).toBe('| x |  |');
  });

  it('flattens merged ranges: top-left keeps the value, others are empty', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'h'); setCell(ws, 1, 2, 'k');
    mergeCells(ws, 'A2:B2');
    setCell(ws, 2, 1, 'merged'); // top-left of the merge
    const md = worksheetToMarkdownTable(ws, 'A1:B2');
    const rows = md.split('\n');
    expect(rows[2]).toBe('| merged |  |');
  });

  it('returns header + separator only when the range is one row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b');
    expect(worksheetToMarkdownTable(ws, 'A1:B1')).toBe(['| a | b |', '| --- | --- |'].join('\n'));
  });
});
