// Scenario 02: every formula kind ECMA-376 specifies.
// Output: 02-formulas.xlsx
//
// What to verify in Excel:
// - A1..A3 numbers; B1 = SUM (normal formula). Excel should display "60".
// - C1:C3 array formula {= A1:A3 * 2}. Click any cell → formula bar
//   shows the array formula braces.
// - D1:D5 shared formula derived from D1 = A1 + 1. D2..D5 should each
//   show A2+1 .. A5+1 with relative reference shifting.
// - E1 = IF(A1>0, "positive", "non-positive") with cachedValue.
// - F1 cell with #N/A error returned by NA().

import { describe, expect, it } from 'vitest';
import { setArrayFormula, setFormula, setSharedFormula } from '../../../src/xlsx/cell/index';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { setCell } from '../../../src/xlsx/worksheet/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 02 — formulas', () => {
  it('writes 02-formulas.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Formulas');

    // Source data for formulas to reference.
    setCell(ws, 1, 1, 10);
    setCell(ws, 2, 1, 20);
    setCell(ws, 3, 1, 30);
    setCell(ws, 4, 1, 40);
    setCell(ws, 5, 1, 50);

    // Normal formula with cached value
    setCell(ws, 1, 2, 0);
    const b1 = ws.rows.get(1)?.get(2);
    if (b1) setFormula(b1, 'SUM(A1:A3)', { cachedValue: 60 });

    // Array formula
    setCell(ws, 1, 3, 0);
    const c1 = ws.rows.get(1)?.get(3);
    if (c1) setArrayFormula(c1, 'C1:C3', 'A1:A3*2', { cachedValue: 20 });

    // Shared formula — origin at D1
    setCell(ws, 1, 4, 0);
    const d1 = ws.rows.get(1)?.get(4);
    if (d1) setSharedFormula(d1, /* si */ 0, 'A1+1', 'D1:D5', { cachedValue: 11 });
    // Subsequent references — empty formula text + same si
    for (let r = 2; r <= 5; r++) {
      setCell(ws, r, 4, 0);
      const c = ws.rows.get(r)?.get(4);
      if (c) setSharedFormula(c, 0, '', undefined, { cachedValue: 10 + r * 10 + 1 });
    }

    // IF formula
    setCell(ws, 1, 5, 0);
    const e1 = ws.rows.get(1)?.get(5);
    if (e1) setFormula(e1, 'IF(A1>0,"positive","non-positive")', { cachedValue: 'positive' });

    // Formula returning an error
    setCell(ws, 1, 6, 0);
    const f1 = ws.rows.get(1)?.get(6);
    if (f1) setFormula(f1, 'NA()');

    const result = await writeWorkbook('02-formulas.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
