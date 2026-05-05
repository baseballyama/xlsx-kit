// Formula text round-trip with XML-special characters. `<`, `>`, `&`,
// `"` all need to escape inside the `<f>` text content, and string
// comparisons inside formulas (`IF(A1="x<y", …)`) must survive the
// cycle.

import { describe, expect, it } from 'vitest';
import { type FormulaValue, setFormula } from '../../src/cell/cell';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

const roundTripFormula = async (formula: string): Promise<string | undefined> => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Sheet1');
  setCell(ws, 1, 1, 0); // anchor an editable cell
  const target = ws.rows.get(1)?.get(1);
  if (!target) throw new Error('cell missing');
  setFormula(target, formula);

  const bytes = await workbookToBytes(wb);
  const wb2 = await loadWorkbook(fromBuffer(bytes));
  const ref0 = wb2.sheets[0];
  if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
  const cell = ref0.sheet.rows.get(1)?.get(1);
  const v = cell?.value;
  if (v === null || v === undefined || typeof v !== 'object' || !('kind' in v) || v.kind !== 'formula') return undefined;
  return (v as FormulaValue).formula;
};

describe('phase-3 — formula-text XML escape round-trip', () => {
  it('preserves `<` / `>` / `&` inside the formula body', async () => {
    expect(await roundTripFormula('IF(A1<10, A1, 10)')).toBe('IF(A1<10, A1, 10)');
    expect(await roundTripFormula('IF(A1>10, A1, 10)')).toBe('IF(A1>10, A1, 10)');
    expect(await roundTripFormula('A1&" rows"')).toBe('A1&" rows"');
  });

  it('preserves quoted strings inside a formula', async () => {
    expect(await roundTripFormula('IF(A1="x<y", "less", "ge")')).toBe('IF(A1="x<y", "less", "ge")');
    expect(await roundTripFormula('CONCAT("a&b", "c<d>e")')).toBe('CONCAT("a&b", "c<d>e")');
  });

  it('preserves a long realistic formula with nested calls', async () => {
    const f = 'IFERROR(VLOOKUP(A1&"_"&B1, Lookup!$A$1:$Z$100, 5, FALSE), "n/a")';
    expect(await roundTripFormula(f)).toBe(f);
  });
});
