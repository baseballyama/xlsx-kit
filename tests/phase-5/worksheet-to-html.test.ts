// Tests for worksheetToHtml — range → <table> rendering.

import { describe, expect, it } from 'vitest';
import { setBold } from '../../src/styles/cell-style';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { worksheetToHtml } from '../../src/worksheet/html';
import { mergeCells, setCell } from '../../src/worksheet/worksheet';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('worksheetToHtml', () => {
  it('renders a plain range as a basic <table>', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    expect(worksheetToHtml(wb, ws, 'A1:B2')).toBe(
      '<table><tr><td>name</td><td>age</td></tr><tr><td>Alice</td><td>30</td></tr></table>',
    );
  });

  it('emits inline style attributes for styled cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'bold');
    setBold(wb, cellAt(ws, 1, 1));
    const html = worksheetToHtml(wb, ws, 'A1');
    expect(html).toContain('font-weight: bold');
    expect(html).toContain('<td style="');
  });

  it('collapses merged ranges with rowspan / colspan and skips covered slots', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'header');
    mergeCells(ws, 'A1:B2');
    setCell(ws, 3, 1, 'a'); setCell(ws, 3, 2, 'b');
    const html = worksheetToHtml(wb, ws, 'A1:B3');
    // top-left cell of the merge has rowspan=2, colspan=2; the other
    // three slots in the merged area produce no <td>.
    expect(html).toContain('rowspan="2"');
    expect(html).toContain('colspan="2"');
    // First row should have one cell; second row should have zero;
    // third row should have two cells.
    const rows = html.match(/<tr>(.*?)<\/tr>/g) ?? [];
    expect(rows.length).toBe(3);
    expect((rows[0]?.match(/<td/g) ?? []).length).toBe(1);
    expect((rows[1]?.match(/<td/g) ?? []).length).toBe(0);
    expect((rows[2]?.match(/<td/g) ?? []).length).toBe(2);
  });

  it('renders empty cells as empty <td>', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    // (1,2) intentionally empty
    expect(worksheetToHtml(wb, ws, 'A1:B1')).toBe('<table><tr><td>a</td><td></td></tr></table>');
  });

  it('HTML-escapes cell values containing <, >, &, "', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, '<b>"AT&T"</b>');
    const html = worksheetToHtml(wb, ws, 'A1');
    expect(html).toContain('&lt;b&gt;&quot;AT&amp;T&quot;&lt;/b&gt;');
    expect(html).not.toContain('<b>');
  });

  it('only includes cells inside the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'in');
    setCell(ws, 5, 5, 'out-of-range');
    const html = worksheetToHtml(wb, ws, 'A1:B2');
    expect(html).toContain('in');
    expect(html).not.toContain('out-of-range');
  });
});
