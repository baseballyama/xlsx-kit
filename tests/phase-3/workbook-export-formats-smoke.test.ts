// Smoke test: full export-format matrix (CSV / HTML / Markdown / Text)
// against a single styled, multi-sheet workbook. Verifies the four
// workbook-wide Record helpers all produce non-empty output for
// populated worksheets and "" for empty ones.

import { describe, expect, it } from 'vitest';
import { setBold, setCellBackgroundColor } from '../../src/styles/cell-style';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  describeWorkbook,
  getWorkbookAsCsvBundle,
  getWorkbookAsCsvRecord,
  getWorkbookAsHtmlRecord,
  getWorkbookAsMarkdownRecord,
  getWorkbookAsTextTableRecord,
} from '../../src/workbook/workbook';
import { mergeCells, setCell } from '../../src/worksheet/worksheet';
import { unzipSync } from 'fflate';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('workbook export-format smoke', () => {
  it('runs the full CSV/HTML/Markdown/Text matrix on a styled multi-sheet workbook', () => {
    const wb = createWorkbook();

    // Sheet "People" with styled header + a small dataset + a merge.
    const people = addWorksheet(wb, 'People');
    setCell(people, 1, 1, 'name'); setCell(people, 1, 2, 'age');
    setCell(people, 2, 1, 'Alice'); setCell(people, 2, 2, 30);
    setCell(people, 3, 1, 'Bob'); setCell(people, 3, 2, 25);
    setBold(wb, cellAt(people, 1, 1));
    setCellBackgroundColor(wb, cellAt(people, 1, 2), 'FFFF00');
    mergeCells(people, 'A4:B4');
    setCell(people, 4, 1, 'totals row');

    // Sheet "Stock" with values + special characters.
    const stock = addWorksheet(wb, 'Stock');
    setCell(stock, 1, 1, 'symbol'); setCell(stock, 1, 2, 'note');
    setCell(stock, 2, 1, 'AT&T'); setCell(stock, 2, 2, 'has | pipe');
    setCell(stock, 3, 1, 'X<>Y'); setCell(stock, 3, 2, 'multi\nline');

    // Empty worksheet.
    addWorksheet(wb, 'Empty');

    // Chartsheet (must be skipped by every record).
    addChartsheet(wb, 'Chart1');

    // Walk the four formats.
    const csv = getWorkbookAsCsvRecord(wb);
    const html = getWorkbookAsHtmlRecord(wb);
    const md = getWorkbookAsMarkdownRecord(wb);
    const txt = getWorkbookAsTextTableRecord(wb);

    // All four records have the same key set (worksheets only, no chartsheet).
    const keys = ['Empty', 'People', 'Stock'];
    expect(Object.keys(csv).sort()).toEqual(keys);
    expect(Object.keys(html).sort()).toEqual(keys);
    expect(Object.keys(md).sort()).toEqual(keys);
    expect(Object.keys(txt).sort()).toEqual(keys);

    // Empty sheet → empty string in every format.
    for (const r of [csv, html, md, txt]) {
      expect(r['Empty']).toBe('');
    }

    // Populated sheets: every format produces non-empty output, with
    // format-specific markers we can spot-check.
    expect(csv['People']).toContain('name,age\n');
    expect(html['People']).toContain('<table>');
    expect(html['People']).toContain('font-weight: bold');
    expect(md['People']).toContain('| name | age |');
    expect(md['People']).toContain('| --- | --- |');
    expect(txt['People']).toMatch(/^\| name +\| age +\|/);
    expect(txt['People']).toContain('+');

    // Special chars are escaped distinctly per format.
    expect(csv['Stock']).toContain('AT&T'); // CSV doesn't HTML-escape &
    expect(html['Stock']).toContain('AT&amp;T');
    expect(md['Stock']).toContain('AT&T');
    expect(md['Stock']).toContain('has \\| pipe'); // pipe escaped in markdown
    expect(txt['Stock']).toContain('AT&T');

    // CSV bundle round-trip: the zip contains a .csv per worksheet.
    const bundle = getWorkbookAsCsvBundle(wb);
    const entries = unzipSync(bundle);
    expect(Object.keys(entries).sort()).toEqual(['Empty.csv', 'People.csv', 'Stock.csv']);

    // describeWorkbook is in the same export family — sanity check it works too.
    const overview = describeWorkbook(wb);
    expect(overview.worksheetCount).toBe(3);
    expect(overview.chartsheetCount).toBe(1);
    expect(overview.cellCount).toBeGreaterThan(0);
  });
});
