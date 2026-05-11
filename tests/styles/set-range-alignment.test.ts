// Tests for setRangeAlignment — bulk Alignment setter (merge / replace).

import { describe, expect, it } from 'vitest';
import {
  alignCellHorizontal,
  getCellAlignment,
  indentCell,
  setRangeAlignment,
} from '../../src/styles/cell-style';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('setRangeAlignment', () => {
  it("merge mode (default) overlays partial onto each cell's existing alignment", () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    indentCell(wb, cellAt(ws, 1, 1), 3); // pre-existing indent
    setRangeAlignment(wb, ws, 'A1', { horizontal: 'right' });
    const a = getCellAlignment(wb, cellAt(ws, 1, 1));
    expect(a.horizontal).toBe('right');
    // indent survives because merge mode is in effect
    expect(a.indent).toBe(3);
  });

  it('replace mode wipes axes that the patch did not include', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    indentCell(wb, cellAt(ws, 1, 1), 3);
    setRangeAlignment(wb, ws, 'A1', { horizontal: 'right' }, 'replace');
    const a = getCellAlignment(wb, cellAt(ws, 1, 1));
    expect(a.horizontal).toBe('right');
    // indent is gone in replace mode
    expect(a.indent).toBeUndefined();
  });

  it('merge preserves a pre-existing horizontal when only vertical is supplied', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    alignCellHorizontal(wb, cellAt(ws, 1, 1), 'left');
    setRangeAlignment(wb, ws, 'A1', { vertical: 'center' });
    const a = getCellAlignment(wb, cellAt(ws, 1, 1));
    expect(a.horizontal).toBe('left');
    expect(a.vertical).toBe('center');
  });

  it('walks the whole range — every cell receives the patch', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'c');
    setCell(ws, 2, 2, 'd');
    setRangeAlignment(wb, ws, 'A1:B2', { horizontal: 'center' });
    for (const [row, col] of [[1, 1], [1, 2], [2, 1], [2, 2]] as const) {
      expect(getCellAlignment(wb, cellAt(ws, row, col)).horizontal).toBe('center');
    }
  });

  it('materialises empty cells so the patch is observable', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeAlignment(wb, ws, 'A1:B2', { wrapText: true });
    for (const [row, col] of [[1, 1], [1, 2], [2, 1], [2, 2]] as const) {
      expect(getCellAlignment(wb, cellAt(ws, row, col)).wrapText).toBe(true);
    }
  });
});
