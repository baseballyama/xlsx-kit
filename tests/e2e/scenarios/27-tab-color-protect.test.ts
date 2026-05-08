// Scenario 27: per-sheet tab colors + sheet protection (read-only).
// Output: 27-tab-color-protect.xlsx
//
// What to verify in Excel:
// - Three sheet tabs at the bottom show distinct colors:
//   * "Red"   tab → solid red (RGB FFFF0000)
//   * "Green" tab → solid green (RGB FF00B050)
//   * "Blue"  tab → solid blue (RGB FF0070C0, theme accent 1 style)
// - "Locked" sheet is **protected**: typing into a cell shows
//   "The cell or chart you are trying to change is on a protected
//   sheet" — Review → Unprotect Sheet (no password) lifts it.
// - Other sheets are editable normally.
//
// Both tabColor and sheetProtection now go through typed APIs:
// `sheetProperties.tabColor` (B7) and `sheetProtection` (B5).

import { describe, expect, it } from 'vitest';
import { makeColor } from '../../../src/styles/index';
import { addWorksheet, createWorkbook } from '../../../src/workbook/index';
import { makeSheetProperties, makeSheetProtection, setCell } from '../../../src/worksheet/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 27 — tab color + sheet protection', () => {
  it('writes 27-tab-color-protect.xlsx', async () => {
    const wb = createWorkbook();

    const setTabColor = (ws: Parameters<typeof setCell>[0], rgb: string): void => {
      ws.sheetProperties = makeSheetProperties({ tabColor: makeColor({ rgb }) });
    };

    const red = addWorksheet(wb, 'Red');
    setCell(red, 1, 1, 'tab color: FFFF0000 (red)');
    setTabColor(red, 'FFFF0000');

    const green = addWorksheet(wb, 'Green');
    setCell(green, 1, 1, 'tab color: FF00B050 (green)');
    setTabColor(green, 'FF00B050');

    const blue = addWorksheet(wb, 'Blue');
    setCell(blue, 1, 1, 'tab color: FF0070C0 (blue)');
    setTabColor(blue, 'FF0070C0');

    const locked = addWorksheet(wb, 'Locked');
    setCell(locked, 1, 1, 'This sheet is protected.');
    setCell(locked, 2, 1, 'Try editing me — Excel will refuse.');
    setCell(locked, 3, 1, 'Review → Unprotect Sheet to lift it.');
    locked.sheetProtection = makeSheetProtection({
      sheet: true,
      objects: true,
      scenarios: true,
      formatCells: false,
      formatColumns: false,
      formatRows: false,
      insertColumns: false,
      insertRows: false,
      insertHyperlinks: false,
      deleteColumns: false,
      deleteRows: false,
      selectLockedCells: false,
      selectUnlockedCells: false,
      sort: false,
      autoFilter: false,
      pivotTables: false,
    });

    const result = await writeWorkbook('27-tab-color-protect.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
