// Scenario 25: alignment edge cases + advanced number formats.
// Output: 25-alignment-numfmt.xlsx
//
// What to verify in Excel:
// - Sheet "Align" demonstrates every alignment axis on cells C2..C9:
//   horizontal (left/center/right/fill/justify), vertical (top/center/
//   bottom), wrapText (long string wrapping inside the cell), indent
//   (3 levels), shrinkToFit (long text auto-shrinks), text rotation
//   (45° / 90° / -45° / vertical-stacked 255).
// - Row 5 (the "wrap" row) is 60pt tall so the wrapped text fits.
// - Sheet "NumFmt" shows column B with one cell per format code:
//     0.00 / #,##0 / #,##0.00 / 0% / 0.00% / 0.00E+00 / "$"#,##0.00 /
//     "(\"$\"#,##0.00);[Red](\"$\"#,##0.00);\"-\""  (positive;negative;zero) /
//     [h]:mm:ss / m/d/yyyy / @"@" (text passthrough) /
//     "# ?/?" / "# ??/??" (fractions).

import { describe, expect, it } from 'vitest';
import { setCellAlignment, setCellNumberFormat } from '../../../src/xlsx/styles/index';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { setCell, setColumnWidth, setRowHeight } from '../../../src/xlsx/worksheet/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 25 — alignment / advanced number formats', () => {
  it('writes 25-alignment-numfmt.xlsx', async () => {
    const wb = createWorkbook();

    // Sheet 1: alignment showcase
    const align = addWorksheet(wb, 'Align');
    setColumnWidth(align, 1, 18); // labels
    setColumnWidth(align, 3, 28); // alignment cells (need width to see effects)
    setRowHeight(align, 5, 60); // wrapText row

    const cases: Array<{ row: number; label: string; value: string; alignment: Parameters<typeof setCellAlignment>[2] }> = [
      { row: 2, label: 'horizontal: left', value: 'left', alignment: { horizontal: 'left' } },
      { row: 3, label: 'horizontal: center', value: 'center', alignment: { horizontal: 'center', vertical: 'center' } },
      { row: 4, label: 'horizontal: right', value: 'right', alignment: { horizontal: 'right' } },
      {
        row: 5,
        label: 'wrapText',
        value: 'This is a long line that should wrap inside the cell when wrapText is on.',
        alignment: { horizontal: 'left', vertical: 'top', wrapText: true },
      },
      { row: 6, label: 'indent: 3', value: 'indent', alignment: { horizontal: 'left', indent: 3 } },
      {
        row: 7,
        label: 'shrinkToFit',
        value: 'shrink-this-very-long-string-please-it-should-fit',
        alignment: { horizontal: 'left', shrinkToFit: true },
      },
      { row: 8, label: 'rotation 45°', value: 'rot-45', alignment: { horizontal: 'center', textRotation: 45 } },
      { row: 9, label: 'rotation 90°', value: 'rot-90', alignment: { horizontal: 'center', textRotation: 90 } },
      { row: 10, label: 'rotation -45 (=135)', value: 'rot-neg45', alignment: { horizontal: 'center', textRotation: 135 } },
      { row: 11, label: 'vertical stacked (255)', value: 'stacked', alignment: { horizontal: 'center', textRotation: 255 } },
    ];

    setCell(align, 1, 1, 'Property');
    setCell(align, 1, 3, 'Cell');
    cases.forEach(({ row, label, value, alignment }) => {
      setCell(align, row, 1, label);
      const cell = setCell(align, row, 3, value);
      setCellAlignment(wb, cell, alignment);
    });
    // Make rotation rows tall enough that the rotated text fits.
    [8, 9, 10, 11].forEach((r) => setRowHeight(align, r, 80));

    // Sheet 2: number format showcase
    const numfmt = addWorksheet(wb, 'NumFmt');
    setColumnWidth(numfmt, 1, 32); // format-code label
    setColumnWidth(numfmt, 2, 22); // formatted value

    const nfCases: Array<{ row: number; code: string; value: string | number }> = [
      { row: 1, code: '0.00', value: 1234.5 },
      { row: 2, code: '#,##0', value: 1234567 },
      { row: 3, code: '#,##0.00', value: 1234567.89 },
      { row: 4, code: '0%', value: 0.42 },
      { row: 5, code: '0.00%', value: 0.4275 },
      { row: 6, code: '0.00E+00', value: 1234567 },
      { row: 7, code: '"$"#,##0.00', value: 99.5 },
      { row: 8, code: '"$"#,##0.00;[Red]"$"#,##0.00;"-"', value: -50 },
      { row: 9, code: '"$"#,##0.00;[Red]"$"#,##0.00;"-"', value: 0 },
      { row: 10, code: '[h]:mm:ss', value: 1.5 }, // 1.5 days = 36:00:00
      { row: 11, code: 'm/d/yyyy', value: 45000 }, // some serial date
      { row: 12, code: '@', value: 'text passthrough' },
      { row: 13, code: '# ?/?', value: 1.5 }, // 1 1/2
      { row: 14, code: '# ??/??', value: 0.123 }, // ~ 1/8
    ];

    setCell(numfmt, 1, 1, 'format code');
    setCell(numfmt, 1, 2, 'value');
    nfCases.forEach(({ row, code, value }) => {
      const labelRow = row + 1; // skip header
      setCell(numfmt, labelRow, 1, code);
      const cell = setCell(numfmt, labelRow, 2, value);
      setCellNumberFormat(wb, cell, code);
    });

    const result = await writeWorkbook('25-alignment-numfmt.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
