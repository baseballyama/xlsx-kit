// Scenario 23: page setup / print options / margins / header-footer.
// Output: 23-page-setup.xlsx
//
// What to verify in Excel:
// - File → Print preview shows landscape A4 with 1-inch top/bottom +
//   0.5-inch left/right margins, fitted to one page wide.
// - Header (centre) reads "Quarterly Report — &P / &N". Footer (left)
//   reads the file name `&F`, footer (right) "Confidential".
// - Sheet has 80 rows so the print preview spans multiple pages, with
//   a manual page break sitting above row 41 (View → Page Break Preview
//   shows the dashed line at the break).
//
// Wired through the typed `printOptions` / `pageMargins` / `pageSetup` /
// `headerFooter` / `rowBreaks` APIs (B6 in docs/plan/13).

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { makeHeaderFooter, makePageMargins, makePageSetup, makePrintOptions, setCell } from '../../../src/xlsx/worksheet/index';
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

    ws.printOptions = makePrintOptions({ horizontalCentered: true, gridLines: true });
    ws.pageMargins = makePageMargins({ left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.3, footer: 0.3 });
    ws.pageSetup = makePageSetup({
      paperSize: 9, // A4
      orientation: 'landscape',
      fitToWidth: 1,
      fitToHeight: 0,
      horizontalDpi: 300,
      verticalDpi: 300,
    });
    ws.headerFooter = makeHeaderFooter({
      differentFirst: false,
      differentOddEven: false,
      oddHeader: '&LQuarterly&CQuarterly Report — &P / &N&R&D',
      oddFooter: '&L&F&CPage &P of &N&RConfidential',
    });

    // Manual page break above row 41 — Page Break Preview shows the line.
    ws.rowBreaks.push({ id: 40, max: 16383, man: true });

    const result = await writeWorkbook('23-page-setup.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
