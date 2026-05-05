// Scenario 15: defined names (workbook-scoped + sheet-scoped + print
// area / print titles).
// Output: 15-defined-names.xlsx
//
// What to verify in Excel:
// - Formulas → Name Manager shows the four names: `total`, `tax`,
//   the sheet-scoped `Sheet1!region`, and `_xlnm.Print_Area`.
// - Cell A10 = total uses the named range and shows 1500.
// - File → Print should preview only the print-area band A1:C5.
// - The first row appears at the top of every printed page (print
//   titles).

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook, setCell } from '../../../src/index';
import { makeDefinedName } from '../../../src/workbook/defined-names';
import { writeWorkbook } from '../_helpers';

describe('e2e 15 — defined names (named ranges + print area)', () => {
  it('writes 15-defined-names.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    setCell(ws, 1, 1, 'Region');
    setCell(ws, 1, 2, 'Sales');
    setCell(ws, 1, 3, 'Tax');
    const data: Array<[string, number, number]> = [
      ['West', 500, 50],
      ['East', 700, 70],
      ['North', 300, 30],
    ];
    data.forEach((row, i) => {
      setCell(ws, i + 2, 1, row[0]);
      setCell(ws, i + 2, 2, row[1]);
      setCell(ws, i + 2, 3, row[2]);
    });

    setCell(ws, 10, 1, '=SUM(total)');
    setCell(ws, 11, 1, '=SUM(tax)');

    wb.definedNames.push(
      makeDefinedName({ name: 'total', value: 'Sheet1!$B$2:$B$4' }),
      makeDefinedName({ name: 'tax', value: 'Sheet1!$C$2:$C$4' }),
      makeDefinedName({ name: 'region', value: 'Sheet1!$A$2:$A$4', scope: 0 }),
      makeDefinedName({ name: '_xlnm.Print_Area', value: 'Sheet1!$A$1:$C$5', scope: 0 }),
      makeDefinedName({ name: '_xlnm.Print_Titles', value: "'Sheet1'!$1:$1", scope: 0 }),
    );

    const result = await writeWorkbook('15-defined-names.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
