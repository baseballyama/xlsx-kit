// Tests for worksheetToTextTable — ASCII-art text table renderer.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { worksheetToTextTable } from '../../src/worksheet/text';
import { mergeCells, setCell } from '../../src/worksheet/worksheet';

describe('worksheetToTextTable', () => {
  it('renders a header + data rows separated by a +---+ border, padded to column width', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    setCell(ws, 3, 1, 'Bo'); setCell(ws, 3, 2, 5);
    expect(worksheetToTextTable(ws, 'A1:B3')).toBe(
      [
        '| name  | age |',
        '+-------+-----+',
        '| Alice | 30  |',
        '| Bo    | 5   |',
      ].join('\n'),
    );
  });

  it('handles unequal column widths (each column gets its own width)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'short'); setCell(ws, 1, 2, 'X');
    setCell(ws, 2, 1, 'verylongvalue'); setCell(ws, 2, 2, 'tiny');
    const lines = worksheetToTextTable(ws, 'A1:B2').split('\n');
    // Each row should have the same string length (consistent padding).
    const lengths = new Set(lines.map((l) => l.length));
    expect(lengths.size).toBe(1);
  });

  it('flattens merged ranges (top-left value, others empty)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'h'); setCell(ws, 1, 2, 'k');
    mergeCells(ws, 'A2:B2');
    setCell(ws, 2, 1, 'merged');
    const md = worksheetToTextTable(ws, 'A1:B2');
    // The data row should show 'merged' in the first column and an
    // empty cell in the second (padded to width).
    expect(md).toContain('| merged |');
  });

  it('replaces newlines inside cell values with a single space', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'col');
    setCell(ws, 2, 1, 'a\nb');
    const txt = worksheetToTextTable(ws, 'A1:A2');
    expect(txt).toContain('| a b |');
    expect(txt).not.toContain('\n a b');
  });

  it('returns header-row + separator only when the range is one row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a'); setCell(ws, 1, 2, 'b');
    expect(worksheetToTextTable(ws, 'A1:B1')).toBe(['| a | b |', '+---+---+'].join('\n'));
  });
});
