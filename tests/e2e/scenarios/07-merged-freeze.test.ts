// Scenario 07: merged cells + freeze panes.
// Output: 07-merged-freeze.xlsx
//
// What to verify in Excel:
// - Sheet "Merged":
//   - A1:C1 is one merged cell showing "Header spans 3 cols".
//   - A2:A4 is one merged cell showing "Spans 3 rows" vertically.
//   - E5:G7 is a 3x3 merge with "Block".
// - Sheet "Freeze":
//   - First row + first column are frozen (visible while scrolling).
//   - Cells beyond E5 fill with "row-N col-M" labels so the freeze is
//     verifiable by scrolling right or down.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  mergeCells,
  setCell,
  setCellByCoord,
  setFreezePanes,
} from '../../../src/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 07 — merged cells + freeze panes', () => {
  it('writes 07-merged-freeze.xlsx', async () => {
    const wb = createWorkbook();

    const merged = addWorksheet(wb, 'Merged');
    setCellByCoord(merged, 'A1', 'Header spans 3 cols');
    mergeCells(merged, 'A1:C1');
    setCellByCoord(merged, 'A2', 'Spans 3 rows');
    mergeCells(merged, 'A2:A4');
    setCellByCoord(merged, 'E5', 'Block');
    mergeCells(merged, 'E5:G7');

    const freeze = addWorksheet(wb, 'Freeze');
    setCell(freeze, 1, 1, 'r/c');
    for (let c = 2; c <= 30; c++) setCell(freeze, 1, c, `col-${c}`);
    for (let r = 2; r <= 50; r++) setCell(freeze, r, 1, `row-${r}`);
    for (let r = 2; r <= 50; r++) {
      for (let c = 2; c <= 30; c++) {
        setCell(freeze, r, c, r * 100 + c);
      }
    }
    setFreezePanes(freeze, 'B2');

    const result = await writeWorkbook('07-merged-freeze.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
