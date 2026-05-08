// Scenario 06: 23 built-in named styles from the "Cell Styles" gallery.
// Output: 06-named-styles.xlsx
//
// What to verify in Excel:
// - Each row's column B should look exactly like the gallery preview
//   for that style: Good (green-on-light), Bad (red-on-pink), Title
//   (Cambria 18pt teal), Headline 1..4, Note (yellow), Hyperlink
//   (blue underline), Comma / Currency / Percent number formats.

import { describe, expect, it } from 'vitest';
import { BUILTIN_NAMED_STYLES, applyBuiltinStyle } from '../../../src/xlsx/styles/index';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { setCell } from '../../../src/xlsx/worksheet/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 06 — built-in named styles', () => {
  it('writes 06-named-styles.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Named Styles');

    setCell(ws, 1, 1, 'Style name');
    setCell(ws, 1, 2, 'Sample');

    let r = 2;
    for (const name of Object.keys(BUILTIN_NAMED_STYLES)) {
      setCell(ws, r, 1, name);
      const isNumber = name === 'Comma' || name === 'Comma [0]' || name === 'Currency' || name === 'Currency [0]' || name === 'Percent';
      const sample = setCell(ws, r, 2, isNumber ? 1234.56 : `style: ${name}`);
      applyBuiltinStyle(wb, sample, name);
      r++;
    }

    const result = await writeWorkbook('06-named-styles.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
