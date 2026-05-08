// Tests for iterAllCells.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  iterAllCells,
} from '../../src/xlsx/workbook/workbook';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

describe('iterAllCells', () => {
  it('yields cells from every worksheet in tab-strip order', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'a1');
    setCell(a, 2, 1, 'a2');
    setCell(b, 1, 1, 'b1');
    const seen = [...iterAllCells(wb)].map(({ sheet, cell }) => `${sheet.title}:${cell.value}`);
    expect(seen).toEqual(['A:a1', 'A:a2', 'B:b1']);
  });

  it('skips chartsheets', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    addChartsheet(wb, 'Chart');
    setCell(a, 1, 1, 'a1');
    const seen = [...iterAllCells(wb)];
    expect(seen.length).toBe(1);
    expect(seen[0]?.sheet.title).toBe('A');
  });

  it('within a sheet, iterates row-then-column ascending', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 5, 2, 'r5c2');
    setCell(a, 1, 3, 'r1c3');
    setCell(a, 1, 1, 'r1c1');
    setCell(a, 5, 1, 'r5c1');
    const order = [...iterAllCells(wb)].map(({ cell }) => `${cell.row}:${cell.col}`);
    expect(order).toEqual(['1:1', '1:3', '5:1', '5:2']);
  });

  it('empty workbook → empty iteration', () => {
    const wb = createWorkbook();
    expect([...iterAllCells(wb)]).toEqual([]);
  });
});
