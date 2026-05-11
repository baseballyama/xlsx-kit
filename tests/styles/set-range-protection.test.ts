// Tests for setRangeProtection — bulk Protection setter on a range.

import { describe, expect, it } from 'vitest';
import {
  getCellFont,
  getCellProtection,
  setBold,
  setRangeProtection,
} from '../../../src/styles/cell-style';
import { addWorksheet, createWorkbook } from '../../../src/workbook/workbook';
import { setCell } from '../../../src/worksheet/worksheet';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('setRangeProtection', () => {
  it('applies locked=false to every cell in the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 'b');
    setCell(ws, 2, 1, 'c');
    setCell(ws, 2, 2, 'd');
    setRangeProtection(wb, ws, 'A1:B2', { locked: false });
    for (const [row, col] of [[1, 1], [1, 2], [2, 1], [2, 2]] as const) {
      const p = getCellProtection(wb, cellAt(ws, row, col));
      expect(p.locked).toBe(false);
      expect(p.hidden).toBe(false);
    }
  });

  it('applies hidden=true (locked defaults to false on partial input)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setRangeProtection(wb, ws, 'A1', { hidden: true });
    const p = getCellProtection(wb, cellAt(ws, 1, 1));
    expect(p).toEqual({ locked: false, hidden: true });
  });

  it('coexists with prior style — bold + protection on the same cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'styled');
    setBold(wb, cellAt(ws, 1, 1));
    setRangeProtection(wb, ws, 'A1', { locked: false });
    const p = getCellProtection(wb, cellAt(ws, 1, 1));
    expect(p.locked).toBe(false);
    // bold survives the protection-only patch
    expect(getCellFont(wb, cellAt(ws, 1, 1)).bold).toBe(true);
  });

  it('is a no-op on a range that contains no cells (range walks the bounds anyway)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    // No setCell — range is all-empty.
    expect(() => setRangeProtection(wb, ws, 'A1:B2', { locked: false })).not.toThrow();
    // No cells exist, so protection is read off DEFAULT_PROTECTION.
    setCell(ws, 1, 1, 'late');
    const p = getCellProtection(wb, cellAt(ws, 1, 1));
    // newly-added cell post-call has the range's protection if walking creates cells,
    // OR DEFAULT (locked=true). The current setRangeStyle implementation walks and
    // creates cells, so locked should be false here.
    expect(p.locked).toBe(false);
  });
});
