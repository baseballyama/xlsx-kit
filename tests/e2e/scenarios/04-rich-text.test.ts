// Scenario 04: rich-text inline string in a cell.
// Output: 04-rich-text.xlsx
//
// What to verify in Excel:
// - A1 currently shows the runs concatenated as a plain string
//   (current stage-1 behaviour). Real rich-text fidelity would show
//   the per-run fonts (bold / italic / colored) — that's a known
//   residual.
// - A2 shows the same plain-string result via a non-rich `<t>` cell
//   for comparison.

import { describe, expect, it } from 'vitest';
import { makeRichText, makeTextRun } from '../../../src/xlsx/cell/index';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { setCell } from '../../../src/xlsx/worksheet/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 04 — rich text', () => {
  it('writes 04-rich-text.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Rich');

    const runs = makeRichText([
      makeTextRun('plain '),
      makeTextRun('bold ', { b: true }),
      makeTextRun('italic ', { i: true }),
      makeTextRun('red ', { color: { rgb: 'FFFF0000' } }),
      makeTextRun('14pt', { sz: 14 }),
    ]);
    setCell(ws, 1, 1, { kind: 'rich-text', runs });

    setCell(ws, 2, 1, 'plain bold italic red 14pt');

    const result = await writeWorkbook('04-rich-text.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
