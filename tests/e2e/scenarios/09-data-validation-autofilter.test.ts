// Scenario 09: data validations + autoFilter.
// Output: 09-data-validation.xlsx
//
// What to verify in Excel:
// - Column A "Status" cells (A2:A10) have a dropdown with values
//   "Open" / "In Progress" / "Done". Selecting another value should
//   show the warning "Choose Open / In Progress / Done".
// - Column B "Score" cells (B2:B10) accept only integers 0..100;
//   typing 200 shows the error "Score must be 0..100".
// - The header row A1:C1 has an AutoFilter dropdown — clicking
//   filters the data below.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { setCell } from '../../../src/xlsx/worksheet/index';
import { parseMultiCellRange, parseRange } from '../../../src/xlsx/worksheet/cell-range';
import { writeWorkbook } from '../_helpers';
void parseRange;

describe('e2e 09 — data validation + autoFilter', () => {
  it('writes 09-data-validation.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Tasks');

    setCell(ws, 1, 1, 'Status');
    setCell(ws, 1, 2, 'Score');
    setCell(ws, 1, 3, 'Owner');
    const data: Array<[string, number, string]> = [
      ['Open', 10, 'alice'],
      ['In Progress', 50, 'bob'],
      ['Done', 100, 'carol'],
      ['Open', 25, 'dave'],
      ['In Progress', 75, 'eve'],
      ['Done', 90, 'frank'],
      ['Open', 60, 'grace'],
      ['Done', 100, 'henry'],
      ['In Progress', 40, 'ivy'],
    ];
    data.forEach((row, i) => {
      setCell(ws, i + 2, 1, row[0]);
      setCell(ws, i + 2, 2, row[1]);
      setCell(ws, i + 2, 3, row[2]);
    });

    ws.dataValidations.push({
      type: 'list',
      sqref: parseMultiCellRange('A2:A10'),
      formula1: '"Open,In Progress,Done"',
      showErrorMessage: true,
      errorTitle: 'Invalid status',
      error: 'Choose Open / In Progress / Done',
      errorStyle: 'warning',
    });
    ws.dataValidations.push({
      type: 'whole',
      operator: 'between',
      sqref: parseMultiCellRange('B2:B10'),
      formula1: '0',
      formula2: '100',
      showErrorMessage: true,
      errorTitle: 'Out of range',
      error: 'Score must be 0..100',
      errorStyle: 'stop',
    });

    ws.autoFilter = { ref: 'A1:C10', filterColumns: [] };

    const result = await writeWorkbook('09-data-validation.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
