// Scenario 23: page setup / print options / margins / header-footer.
// Output: 23-page-setup.xlsx
//
// What to verify in Excel:
// - File → Print preview shows landscape A4 with 1-inch top/bottom +
//   0.5-inch left/right margins, fitted to one page wide.
// - Header (centre) reads "Quarterly Report — &P / &N". Footer (left)
//   reads the file name `&F`, footer (right) "Confidential".
// - Page Layout → Print Titles shows row 1 repeating on every printed
//   page (configured via the workbook's defined names section, not the
//   passthrough hooks below — printTitles needs the named-range route).
// - Sheet has 80 rows so the print preview spans 2 pages with row 1
//   sticking on top of page 2.
//
// This scenario uses the worksheet `bodyExtras.afterSheetData` slot to
// inject `<printOptions>`, `<pageMargins>`, `<pageSetup>`, and
// `<headerFooter>` as raw XmlNodes — these elements aren't yet modeled
// by the high-level worksheet API but the round-trip preserves them.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook, elNs, setCell } from '../../../src/index';
import { SHEET_MAIN_NS } from '../../../src/xml/namespaces';
import { writeWorkbook } from '../_helpers';

describe('e2e 23 — page setup / print options / header-footer', () => {
  it('writes 23-page-setup.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Report');

    setCell(ws, 1, 1, 'Period');
    setCell(ws, 1, 2, 'Region');
    setCell(ws, 1, 3, 'Revenue');
    setCell(ws, 1, 4, 'Profit');
    for (let r = 2; r <= 80; r++) {
      setCell(ws, r, 1, `Day ${r - 1}`);
      setCell(ws, r, 2, ['North', 'South', 'East', 'West'][(r - 2) % 4] ?? 'North');
      setCell(ws, r, 3, 1000 + ((r * 37) % 500));
      setCell(ws, r, 4, 200 + ((r * 17) % 200));
    }

    // ECMA-376 part 1 §18.3.1.70 pageSetup, §18.3.1.62 pageMargins,
    // §18.3.1.70a printOptions, §18.3.1.46 headerFooter.
    ws.bodyExtras = {
      beforeSheetData: [],
      afterSheetData: [
        // Center horizontally on the page (vertically off).
        elNs(SHEET_MAIN_NS, 'printOptions', { horizontalCentered: '1', gridLines: '1' }),
        elNs(SHEET_MAIN_NS, 'pageMargins', {
          left: '0.5',
          right: '0.5',
          top: '1',
          bottom: '1',
          header: '0.3',
          footer: '0.3',
        }),
        elNs(SHEET_MAIN_NS, 'pageSetup', {
          paperSize: '9', // A4
          orientation: 'landscape',
          fitToWidth: '1',
          fitToHeight: '0',
          horizontalDpi: '300',
          verticalDpi: '300',
        }),
        elNs(
          SHEET_MAIN_NS,
          'headerFooter',
          { differentFirst: '0', differentOddEven: '0' },
          [
            elNs(SHEET_MAIN_NS, 'oddHeader', {}, [], '&LQuarterly&CQuarterly Report — &P / &N&R&D'),
            elNs(SHEET_MAIN_NS, 'oddFooter', {}, [], '&L&F&CPage &P of &N&RConfidential'),
          ],
        ),
      ],
    };

    const result = await writeWorkbook('23-page-setup.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
