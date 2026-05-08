// Tests for cellRangeFromCells — Cell[] → bounding A1 range.

import { describe, expect, it } from 'vitest';
import { cellRangeFromCells } from '../../src/xlsx/worksheet/cell-range';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

describe('cellRangeFromCells', () => {
  it('returns a single-cell ref for a one-cell input', () => {
    expect(cellRangeFromCells([{ row: 1, col: 1 }])).toBe('A1');
    expect(cellRangeFromCells([{ row: 5, col: 7 }])).toBe('G5');
  });

  it('computes the bounding rectangle for a four-cell rectangle', () => {
    expect(
      cellRangeFromCells([
        { row: 1, col: 1 },
        { row: 1, col: 2 },
        { row: 2, col: 1 },
        { row: 2, col: 2 },
      ]),
    ).toBe('A1:B2');
  });

  it('uses min/max for sparse / disjoint cells (no contiguity check)', () => {
    // Three corners of a 5×3 region.
    expect(
      cellRangeFromCells([
        { row: 1, col: 1 },
        { row: 5, col: 3 },
        { row: 3, col: 2 },
      ]),
    ).toBe('A1:C5');
  });

  it('returns a single-column range for collinear column inputs', () => {
    expect(
      cellRangeFromCells([
        { row: 2, col: 4 },
        { row: 5, col: 4 },
        { row: 7, col: 4 },
      ]),
    ).toBe('D2:D7');
  });

  it('throws on empty input', () => {
    expect(() => cellRangeFromCells([])).toThrow(/non-empty/);
  });

  it('works with real Cell objects from a worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const a = setCell(ws, 1, 1, 'a');
    const b = setCell(ws, 3, 5, 'b');
    expect(cellRangeFromCells([a, b])).toBe('A1:E3');
  });
});
