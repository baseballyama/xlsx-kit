// Scenario 26: two Excel tables on a single worksheet, side-by-side,
// using different table styles. Output: 26-multi-table.xlsx
//
// What to verify in Excel:
// - One sheet "Reports" with two named tables:
//   * `tblOrders` covering A1:D7 with TableStyleLight9 (blue accent),
//     showRowStripes ON.
//   * `tblPayroll` covering F1:I7 with TableStyleMedium14 (green
//     accent), showColumnStripes ON.
// - Both tables have AutoFilter dropdowns on the header row.
// - Type `=SUM(tblOrders[Total])` and `=SUM(tblPayroll[Net])` in a
//   blank cell to verify both structured-reference autocompletes work.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { addTable, setCell, setColumnWidth } from '../../../src/xlsx/worksheet/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 26 — multi-table sheet', () => {
  it('writes 26-multi-table.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Reports');

    // Table 1: tblOrders @ A1:D7
    const ordersHeaders = ['Order', 'Region', 'Qty', 'Total'];
    ordersHeaders.forEach((h, i) => setCell(ws, 1, i + 1, h));
    const ordersData: Array<[string, string, number, number]> = [
      ['ORD-001', 'North', 12, 1450],
      ['ORD-002', 'South', 8, 980],
      ['ORD-003', 'East', 20, 2210],
      ['ORD-004', 'West', 5, 580],
      ['ORD-005', 'North', 15, 1820],
      ['ORD-006', 'East', 9, 1100],
    ];
    ordersData.forEach((row, ri) => row.forEach((v, ci) => setCell(ws, ri + 2, ci + 1, v)));

    addTable(ws, {
      id: 1,
      name: 'tblOrders',
      displayName: 'tblOrders',
      ref: 'A1:D7',
      headerRowCount: 1,
      columns: ordersHeaders.map((name, idx) => ({ id: idx + 1, name })),
      styleInfo: {
        name: 'TableStyleLight9',
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    });

    // Table 2: tblPayroll @ F1:I7
    const payrollHeaders = ['Employee', 'Dept', 'Gross', 'Net'];
    payrollHeaders.forEach((h, i) => setCell(ws, 1, i + 6, h));
    const payrollData: Array<[string, string, number, number]> = [
      ['Alice', 'Eng', 8200, 6500],
      ['Bob', 'Eng', 7800, 6200],
      ['Carol', 'Sales', 7400, 5950],
      ['Dan', 'Ops', 6900, 5500],
      ['Eve', 'Eng', 9100, 7100],
      ['Frank', 'Sales', 7100, 5680],
    ];
    payrollData.forEach((row, ri) => row.forEach((v, ci) => setCell(ws, ri + 2, ci + 6, v)));

    addTable(ws, {
      id: 2,
      name: 'tblPayroll',
      displayName: 'tblPayroll',
      ref: 'F1:I7',
      headerRowCount: 1,
      columns: payrollHeaders.map((name, idx) => ({ id: idx + 1, name })),
      styleInfo: {
        name: 'TableStyleMedium14',
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: false,
        showColumnStripes: true,
      },
    });

    [1, 2, 3, 4, 6, 7, 8, 9].forEach((c) => setColumnWidth(ws, c, 12));

    const result = await writeWorkbook('26-multi-table.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
