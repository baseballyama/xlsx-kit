// Scenario 30: classic single-value formulas. Output: 30-modern-formulas.xlsx
//
// The original scenario covered Excel 365 / 2021 modern formulas (LET,
// LAMBDA, FILTER, SORT, UNIQUE, SEQUENCE, XLOOKUP, BYROW). Two layers
// of Excel constraints make those impractical to ship from this writer
// without a full metadata implementation:
//   1. Bare `LET(...)` / `XLOOKUP(...)` resolve only on Excel 365
//      (LET) or Excel 2019+ (XLOOKUP); older Excel shows #NAME?.
//   2. The `_xlfn.LET` / `_xlfn._xlws.FILTER` storage form Excel writes
//      itself requires a paired `xl/metadata.xml` part + `cm="N"`
//      cell-metadata reference — without those, Excel raises the
//      "We found a problem" recovery dialog on open.
//
// Until the metadata-emitter lands, fall back to formulas that exist in
// every Excel since 2007: AVERAGE / VLOOKUP. Same demonstrative shape
// (single-value arithmetic + lookup), guaranteed to evaluate everywhere.
//
// What to verify in Excel:
// - E2 = AVERAGE(C2:C8) returns the average salary (84428.57).
// - F2 = VLOOKUP("Bob",A2:C8,3,FALSE) returns Bob's salary (88000).

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../../src/workbook/index';
import { setCell } from '../../../src/worksheet/index';
import { setFormula } from '../../../src/cell/cell';
import { writeWorkbook } from '../_helpers';

describe('e2e 30 — single-value formulas', () => {
  it('writes 30-modern-formulas.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Modern');

    // Source table A1:C8
    setCell(ws, 1, 1, 'Name');
    setCell(ws, 1, 2, 'Dept');
    setCell(ws, 1, 3, 'Salary');
    const rows: Array<[string, string, number]> = [
      ['Alice', 'Eng', 95000],
      ['Bob', 'Eng', 88000],
      ['Carol', 'Sales', 79000],
      ['Dan', 'Ops', 71000],
      ['Eve', 'Eng', 102000],
      ['Frank', 'Sales', 76000],
      ['Grace', 'Ops', 80000],
    ];
    rows.forEach((r, i) => r.forEach((v, c) => setCell(ws, i + 2, c + 1, v)));

    setCell(ws, 1, 5, 'AVERAGE');
    setFormula(setCell(ws, 2, 5, ''), 'AVERAGE(C2:C8)', { cachedValue: 84428.57 });

    setCell(ws, 1, 6, 'VLOOKUP');
    setFormula(setCell(ws, 2, 6, ''), 'VLOOKUP("Bob",A2:C8,3,FALSE)', { cachedValue: 88000 });

    const result = await writeWorkbook('30-modern-formulas.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
