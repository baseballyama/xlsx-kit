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
// Tab colors are now wired through the typed `sheetProperties.tabColor`
// API (B7 in docs/plan/13). Sheet protection is still emitted as a raw
// `<sheetProtection>` XmlNode via `bodyExtras.afterSheetData` since the
// high-level worksheet API doesn't yet model sheetProtection.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  elNs,
  makeColor,
  makeSheetProperties,
  setCell,
} from '../../../src/index';
import { SHEET_MAIN_NS } from '../../../src/xml/namespaces';
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

    // Locked sheet — sheet protection (no password) emitted as
    // <sheetProtection sheet="1" objects="1" scenarios="1"/>.
    const locked = addWorksheet(wb, 'Locked');
    setCell(locked, 1, 1, 'This sheet is protected.');
    setCell(locked, 2, 1, 'Try editing me — Excel will refuse.');
    setCell(locked, 3, 1, 'Review → Unprotect Sheet to lift it.');
    locked.bodyExtras = {
      beforeSheetData: [],
      afterSheetData: [
        // sheetProtection lives between mergeCells and pageMargins per
        // ECMA-376 §18.3.1.85, comfortably inside afterSheetData.
        elNs(SHEET_MAIN_NS, 'sheetProtection', {
          sheet: '1',
          objects: '1',
          scenarios: '1',
          formatCells: '0',
          formatColumns: '0',
          formatRows: '0',
          insertColumns: '0',
          insertRows: '0',
          insertHyperlinks: '0',
          deleteColumns: '0',
          deleteRows: '0',
          selectLockedCells: '0',
          sort: '0',
          autoFilter: '0',
          pivotTables: '0',
          selectUnlockedCells: '0',
        }),
      ],
    };

    const result = await writeWorkbook('27-tab-color-protect.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
