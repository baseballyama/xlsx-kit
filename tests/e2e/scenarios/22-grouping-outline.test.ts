// Scenario 22: row + column outline grouping (collapsible groups).
// Output: 22-grouping-outline.xlsx
//
// What to verify in Excel:
// - Outline buttons appear above the column headers and to the left of
//   the row numbers (a tiny "1 / 2" toggle bar). Clicking "1" collapses
//   to summary level, "2" expands to detail level.
// - Rows 3..6 are level-1 grouped (Q1 detail), rows 8..11 are level-1
//   grouped (Q2 detail). Row 7 = Q1 subtotal (visible), row 12 = Q2
//   subtotal, row 13 = grand total.
// - Columns C..D are grouped (level 1) so the user can collapse them
//   leaving column B (Plan) visible next to E (Variance).
// - Column F is hidden — pressing the outline expand button shouldn't
//   reveal it; it requires Format → Hide & Unhide → Unhide Columns.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  hideColumn,
  setCell,
  setColumnDimension,
  setColumnWidth,
  setRowDimension,
  setRowHeight,
} from '../../../src/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 22 — grouping / outline / hide', () => {
  it('writes 22-grouping-outline.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Budget');

    // Header row
    setCell(ws, 1, 1, 'Period');
    setCell(ws, 1, 2, 'Plan');
    setCell(ws, 1, 3, 'Actual Jan');
    setCell(ws, 1, 4, 'Actual Feb');
    setCell(ws, 1, 5, 'Variance');
    setCell(ws, 1, 6, 'Note (hidden col)');

    // Q1 block
    setCell(ws, 2, 1, 'Q1');
    setCell(ws, 3, 1, '  Sales');
    setCell(ws, 3, 2, 1000);
    setCell(ws, 3, 3, 480);
    setCell(ws, 3, 4, 510);
    setCell(ws, 4, 1, '  COGS');
    setCell(ws, 4, 2, -400);
    setCell(ws, 4, 3, -190);
    setCell(ws, 4, 4, -205);
    setCell(ws, 5, 1, '  OpEx');
    setCell(ws, 5, 2, -200);
    setCell(ws, 5, 3, -95);
    setCell(ws, 5, 4, -100);
    setCell(ws, 6, 1, '  Tax');
    setCell(ws, 6, 2, -80);
    setCell(ws, 6, 3, -39);
    setCell(ws, 6, 4, -41);
    setCell(ws, 7, 1, 'Q1 Subtotal');
    setCell(ws, 7, 2, 320);
    setCell(ws, 7, 5, 4);

    // Q2 block
    setCell(ws, 8, 1, '  Sales');
    setCell(ws, 8, 2, 1100);
    setCell(ws, 8, 3, 530);
    setCell(ws, 8, 4, 555);
    setCell(ws, 9, 1, '  COGS');
    setCell(ws, 9, 2, -440);
    setCell(ws, 9, 3, -210);
    setCell(ws, 9, 4, -222);
    setCell(ws, 10, 1, '  OpEx');
    setCell(ws, 10, 2, -210);
    setCell(ws, 10, 3, -100);
    setCell(ws, 10, 4, -106);
    setCell(ws, 11, 1, '  Tax');
    setCell(ws, 11, 2, -85);
    setCell(ws, 11, 3, -41);
    setCell(ws, 11, 4, -43);
    setCell(ws, 12, 1, 'Q2 Subtotal');
    setCell(ws, 12, 2, 365);
    setCell(ws, 12, 5, 3);

    setCell(ws, 13, 1, 'Total');
    setCell(ws, 13, 2, 685);

    // Outline-level rows (Q1 detail rows 3..6, Q2 detail rows 8..11)
    for (const row of [3, 4, 5, 6]) setRowDimension(ws, row, { outlineLevel: 1 });
    for (const row of [8, 9, 10, 11]) setRowDimension(ws, row, { outlineLevel: 1 });
    // Make the subtotal rows taller so they stand out.
    setRowHeight(ws, 7, 22);
    setRowHeight(ws, 12, 22);
    setRowHeight(ws, 13, 26);

    // Custom widths + outline columns C..D, hidden column F.
    setColumnWidth(ws, 1, 22);
    setColumnWidth(ws, 2, 14);
    setColumnDimension(ws, 3, { width: 12, outlineLevel: 1 });
    setColumnDimension(ws, 4, { width: 12, outlineLevel: 1 });
    setColumnWidth(ws, 5, 12);
    hideColumn(ws, 6);

    const result = await writeWorkbook('22-grouping-outline.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
