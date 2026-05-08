// Phase 3 — empty-but-styled cells must keep their styleId.
// `<c r="A1" s="N"/>` (no `<v>` / `<f>`) is how Excel represents a
// cell whose value is empty but whose formatting differs from the
// sheet default. The writer + reader both have to handle this — and
// crucially the writer must NOT skip the cell just because its
// value is null.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addCellXf, defaultCellXf } from '../../src/styles/stylesheet';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

describe('phase-3 — empty-styled cells round-trip', () => {
  it('keeps `<c r="A1" s="N"/>` cells across save → load', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    // Reserve cellXfs[0] for the default so subsequent allocations
    // get slot ≥ 1 and the writer emits an explicit `s="N"` attribute.
    addCellXf(wb.styles, defaultCellXf());
    const styleId = addCellXf(wb.styles, { ...defaultCellXf(), applyFont: true });
    expect(styleId).toBeGreaterThan(0);

    setCell(ws, 1, 1, null, styleId);
    setCell(ws, 1, 2, 'has-value');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');

    const styled = ref0.sheet.rows.get(1)?.get(1);
    expect(styled?.value).toBeNull();
    expect(styled?.styleId).toBe(styleId);

    const valued = ref0.sheet.rows.get(1)?.get(2);
    expect(valued?.value).toBe('has-value');
  });

  it('does not emit empty cells when both value and styleId are default', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    setCell(ws, 1, 1, null, 0);
    setCell(ws, 1, 2, 42);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');

    // Cell A1 had nothing notable — should not be emitted; the
    // reload sees only B1.
    const colMap = ref0.sheet.rows.get(1);
    expect(colMap?.has(1)).toBe(false);
    expect(colMap?.get(2)?.value).toBe(42);
  });

  it('preserves a row of all-empty-but-styled cells', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    addCellXf(wb.styles, defaultCellXf()); // reserve slot 0
    const styleId = addCellXf(wb.styles, { ...defaultCellXf(), applyFill: true });
    setCell(ws, 5, 1, null, styleId);
    setCell(ws, 5, 2, null, styleId);
    setCell(ws, 5, 3, null, styleId);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');

    for (const col of [1, 2, 3]) {
      const cell = ref0.sheet.rows.get(5)?.get(col);
      expect(cell?.styleId).toBe(styleId);
      expect(cell?.value).toBeNull();
    }
  });
});
