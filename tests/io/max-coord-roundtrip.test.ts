// Phase 3 — boundary cells (XFD1048576) round-trip cleanly.
// MAX_ROW = 1_048_576 and MAX_COL = 16_384 (XFD) per ECMA-376 §3.2.
// Pins that the writer / reader handle the largest valid coordinate
// without integer overflow, off-by-one, or ref-string parsing bugs.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { MAX_COL, MAX_ROW } from '../../src/utils/coordinate';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('phase-3 — max-coordinate (XFD1048576) round-trip', () => {
  it('writes and reads back a cell at row=MAX_ROW / col=MAX_COL', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Edge');
    setCell(ws, MAX_ROW, MAX_COL, 'corner');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
    const cell = ref0.sheet.rows.get(MAX_ROW)?.get(MAX_COL);
    expect(cell?.value).toBe('corner');
    expect(cell?.row).toBe(MAX_ROW);
    expect(cell?.col).toBe(MAX_COL);
  });

  it('round-trips a sparse spread across the full sheet (A1, XFD1, A1048576, XFD1048576)', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sparse');
    setCell(ws, 1, 1, 'top-left');
    setCell(ws, 1, MAX_COL, 'top-right');
    setCell(ws, MAX_ROW, 1, 'bottom-left');
    setCell(ws, MAX_ROW, MAX_COL, 'bottom-right');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
    expect(ref0.sheet.rows.get(1)?.get(1)?.value).toBe('top-left');
    expect(ref0.sheet.rows.get(1)?.get(MAX_COL)?.value).toBe('top-right');
    expect(ref0.sheet.rows.get(MAX_ROW)?.get(1)?.value).toBe('bottom-left');
    expect(ref0.sheet.rows.get(MAX_ROW)?.get(MAX_COL)?.value).toBe('bottom-right');
  });
});
