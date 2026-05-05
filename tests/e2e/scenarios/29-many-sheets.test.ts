// Scenario 29: a workbook with many tabs to stress the bottom tab
// strip + cross-sheet references. Output: 29-many-sheets.xlsx
//
// What to verify in Excel:
// - 30 tabs at the bottom. Tabs labelled "M01" through "M30" (months
//   M01..M12 first year, M13..M24 second, M25..M30 third partial).
// - The leading "Summary" tab (sheet 1) shows = 'M01'!A1 / 'M30'!A1
//   etc. — Excel resolves them to "value from M01" / "value from M30".
// - Excel allows tabs to be reordered / hidden via right-click on a
//   tab; verify nothing about saving from this build prevents that.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook, setCell, setColumnWidth } from '../../../src/index';
import { setFormula } from '../../../src/cell/cell';
import { writeWorkbook } from '../_helpers';

describe('e2e 29 — many sheets + cross-sheet references', () => {
  it('writes 29-many-sheets.xlsx', async () => {
    const wb = createWorkbook();

    const summary = addWorksheet(wb, 'Summary');
    setCell(summary, 1, 1, 'Sheet');
    setCell(summary, 1, 2, 'A1 ref');
    setColumnWidth(summary, 1, 14);
    setColumnWidth(summary, 2, 22);

    const N = 30;
    for (let i = 1; i <= N; i++) {
      const name = `M${i.toString().padStart(2, '0')}`;
      const ws = addWorksheet(wb, name);
      setCell(ws, 1, 1, `value from ${name}`);
      setCell(ws, 2, 1, i * 100);

      setCell(summary, i + 1, 1, name);
      const refCell = setCell(summary, i + 1, 2, '');
      setFormula(refCell, `'${name}'!A1`);
    }

    const result = await writeWorkbook('29-many-sheets.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
