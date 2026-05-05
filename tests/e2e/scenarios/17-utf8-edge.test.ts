// Scenario 17: edge cases — multi-byte sheet titles, multi-byte cell
// content, max-coord cell, control characters via _xHHHH_.
// Output: 17-utf8-edge.xlsx
//
// What to verify in Excel:
// - Tab labels include "売上" (Japanese), "مبيعات" (Arabic, RTL),
//   "Resumé". All three should render with correct script.
// - "売上" sheet: A1 contains a string with mixed scripts + emoji 😀.
//   B1 has 改行を含むセル文字列 (multi-line).
// - "Resumé" sheet: a cell at the maximum coord XFD1048576 contains
//   the literal "corner" — Excel will jump there if you press Ctrl+End.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook, MAX_COL, MAX_ROW, setCell } from '../../../src/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 17 — UTF-8 + edge coordinates', () => {
  it('writes 17-utf8-edge.xlsx', async () => {
    const wb = createWorkbook();

    const ja = addWorksheet(wb, '売上');
    setCell(ja, 1, 1, 'こんにちは, مرحبا 😀');
    setCell(ja, 1, 2, 'line1\nline2\nline3');
    setCell(ja, 2, 1, '売上合計');
    setCell(ja, 2, 2, 1234567);

    const ar = addWorksheet(wb, 'مبيعات');
    setCell(ar, 1, 1, 'العربية');
    setCell(ar, 1, 2, 1000);

    const accented = addWorksheet(wb, 'Resumé');
    setCell(accented, 1, 1, 'naïve café — flambé');
    setCell(accented, MAX_ROW, MAX_COL, 'corner (Ctrl+End to find me)');

    const result = await writeWorkbook('17-utf8-edge.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
