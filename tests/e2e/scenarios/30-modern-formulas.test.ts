// Scenario 30: Excel 365 / 2021 modern formulas — LET, LAMBDA,
// FILTER, SORT, UNIQUE, SEQUENCE, XLOOKUP, BYROW, MAP, REDUCE.
// Output: 30-modern-formulas.xlsx
//
// What to verify in Excel:
// - Open in Excel 365 or Excel 2021 — older Excel will show #NAME? for
//   modern functions like LET / LAMBDA / FILTER and that's expected.
// - Each formula should evaluate (Excel auto-recalcs on open) and
//   demonstrate the spilled-array (dynamic array) behavior:
//   * Source data lives in A2:C8 (Name / Dept / Salary).
//   * E2 = LET(...) returning a single weighted average.
//   * G2 = FILTER(A2:C8, B2:B8="Eng") — spills below.
//   * K2 = SORT(A2:A8) — spills below.
//   * M2 = UNIQUE(B2:B8) — spills below.
//   * O2 = SEQUENCE(5) — spills 1..5 down.
//   * Q2 = XLOOKUP("Bob", A2:A8, C2:C8) — single value.
//   * S2 = LAMBDA(x, x*x)(7) — should return 49.
//   * U2 = BYROW(C2:C8, LAMBDA(r, r*1.1)) — spills 7 rows.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook, setCell } from '../../../src/index';
import { setFormula } from '../../../src/cell/cell';
import { writeWorkbook } from '../_helpers';

describe('e2e 30 — modern dynamic-array + lambda formulas', () => {
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

    setCell(ws, 1, 5, 'LET weighted avg');
    setFormula(setCell(ws, 2, 5, ''), 'LET(s,SUM(C2:C8),n,COUNT(C2:C8),s/n)');

    setCell(ws, 1, 7, 'FILTER Eng');
    setFormula(setCell(ws, 2, 7, ''), 'FILTER(A2:C8, B2:B8="Eng")');

    setCell(ws, 1, 11, 'SORT Names');
    setFormula(setCell(ws, 2, 11, ''), 'SORT(A2:A8)');

    setCell(ws, 1, 13, 'UNIQUE Dept');
    setFormula(setCell(ws, 2, 13, ''), 'UNIQUE(B2:B8)');

    setCell(ws, 1, 15, 'SEQUENCE(5)');
    setFormula(setCell(ws, 2, 15, ''), 'SEQUENCE(5)');

    setCell(ws, 1, 17, 'XLOOKUP Bob salary');
    setFormula(setCell(ws, 2, 17, ''), 'XLOOKUP("Bob",A2:A8,C2:C8)');

    setCell(ws, 1, 19, 'LAMBDA(x,x*x)(7)');
    setFormula(setCell(ws, 2, 19, ''), 'LAMBDA(x,x*x)(7)');

    setCell(ws, 1, 21, 'BYROW *1.1');
    setFormula(setCell(ws, 2, 21, ''), 'BYROW(C2:C8,LAMBDA(r,r*1.1))');

    const result = await writeWorkbook('30-modern-formulas.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
