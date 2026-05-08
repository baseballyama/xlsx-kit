// Phase 3 §5.5 — `<f t="dataTable">` round-trip.
// `setDataTableFormula` preserves all dt-specific attributes (ref, r1,
// r2, dt2D, dtr, del1, del2, aca, ca) so Excel "What-if Analysis →
// Data Table" cells survive load → save → load without regressing to
// a normal formula.

import { describe, expect, it } from 'vitest';
import { type CellValue, type FormulaValue, setDataTableFormula } from '../../src/xlsx/cell/cell';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { setCell } from '../../src/xlsx/worksheet/worksheet';

const asFormula = (v: CellValue | undefined): FormulaValue => {
  if (v === null || v === undefined || typeof v !== 'object' || !('kind' in v) || v.kind !== 'formula') {
    throw new Error('cell value is not a formula');
  }
  return v;
};

describe('phase-3 §5.5 — dataTable formula round-trip', () => {
  it('preserves t / ref / r1 / r2 / dt2D across save → load', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    setCell(ws, 1, 1, 10);
    const target = ws.rows.get(1)?.get(1);
    if (!target) throw new Error('cell missing');
    setDataTableFormula(target, 'TABLE(B1,C1)', {
      ref: 'A1:A3',
      r1: '$B$1',
      r2: '$C$1',
      dt2D: true,
      ca: true,
      cachedValue: 100,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
    const cell = ref0.sheet.rows.get(1)?.get(1);
    const f = asFormula(cell?.value);
    expect(f.t).toBe('dataTable');
    expect(f.formula).toBe('TABLE(B1,C1)');
    expect(f.ref).toBe('A1:A3');
    expect(f.r1).toBe('$B$1');
    expect(f.r2).toBe('$C$1');
    expect(f.dt2D).toBe(true);
    expect(f.ca).toBe(true);
    expect(f.cachedValue).toBe(100);
  });

  it('handles a 1-variable column-direction Data Table', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    setCell(ws, 1, 1, 0);
    const target = ws.rows.get(1)?.get(1);
    if (!target) throw new Error('cell missing');
    setDataTableFormula(target, 'TABLE(,A1)', { ref: 'A1:A5', r1: '$A$1' });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
    const cell = ref0.sheet.rows.get(1)?.get(1);
    const f = asFormula(cell?.value);
    expect(f.t).toBe('dataTable');
    expect(f.r1).toBe('$A$1');
    expect(f.r2).toBeUndefined();
    expect(f.dt2D).toBeUndefined();
  });

  it('rejects <f t="dataTable"> without @ref on read', async () => {
    // Manufacture a minimal sheet with a malformed dataTable formula by
    // round-tripping through a workbook → save → load cycle. Simpler:
    // craft the bytes by hand.
    const malformed = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><f t="dataTable">TABLE(B1,)</f></c></row>
  </sheetData>
</worksheet>`;
    const { parseWorksheetXml } = await import('../../src/xlsx/worksheet/reader');
    expect(() =>
      parseWorksheetXml(new TextEncoder().encode(malformed), 'Sheet1', { sharedStrings: [] }),
    ).toThrowError(/missing @ref/);
  });
});
