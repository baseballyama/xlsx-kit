// Scenario 10: an Excel Table (formatted data range with header /
// totals / structured references).
// Output: 10-tables.xlsx
//
// What to verify in Excel:
// - The "Sales" sheet has a table named "tblSales" covering A1:D6
//   with TableStyleMedium2 (or whatever default table style your
//   Excel version applies).
// - Header row + alternating-row banding visible.
// - The Quantity column should support structured references like
//   `=SUM(tblSales[Quantity])` typed into another cell.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook, setCell } from '../../../src/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 10 — Excel table', () => {
  it('writes 10-tables.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sales');

    const headers = ['Product', 'Region', 'Quantity', 'Price'];
    headers.forEach((h, i) => setCell(ws, 1, i + 1, h));

    const data: Array<[string, string, number, number]> = [
      ['Apples', 'West', 100, 1.5],
      ['Apples', 'East', 200, 1.5],
      ['Oranges', 'West', 150, 2.0],
      ['Oranges', 'East', 80, 2.0],
      ['Pears', 'West', 60, 2.5],
    ];
    data.forEach((row, ri) => {
      row.forEach((v, ci) => setCell(ws, ri + 2, ci + 1, v));
    });

    ws.tables.push({
      id: 1,
      name: 'tblSales',
      displayName: 'tblSales',
      ref: 'A1:D6',
      headerRowCount: 1,
      columns: headers.map((name, idx) => ({ id: idx + 1, name })),
      styleInfo: {
        name: 'TableStyleMedium2',
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    });

    const result = await writeWorkbook('10-tables.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
