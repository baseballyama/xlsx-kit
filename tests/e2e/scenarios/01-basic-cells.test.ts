// Scenario 01: every basic cell value type Excel renders natively.
// Output: 01-basic-cells.xlsx
//
// What to verify in Excel:
// - A1..A6 show: 42, 3.14, -1, "hello", "<&>", "with\\ttab\\nnewline\\rcr"
// - B1..B6 show: TRUE, FALSE, the four standard error codes #DIV/0!,
//   #N/A, #REF!, #VALUE!
// - C1..C5 show: blank cells (null values, no styling)
// - D1..D2 show: very large + very small numbers without losing precision

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook, makeErrorValue, setCell } from '../../../src/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 01 — basic cell values', () => {
  it('writes 01-basic-cells.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Basic');

    // Numbers
    setCell(ws, 1, 1, 42);
    setCell(ws, 2, 1, 3.14);
    setCell(ws, 3, 1, -1);
    setCell(ws, 1, 4, 1.7976931348623157e308); // near MAX_VALUE
    setCell(ws, 2, 4, 5e-324); // near MIN_VALUE

    // Strings (sharedStrings + escape)
    setCell(ws, 4, 1, 'hello');
    setCell(ws, 5, 1, '<&>');
    setCell(ws, 6, 1, 'with\ttab\nnewline\rcr');

    // Booleans
    setCell(ws, 1, 2, true);
    setCell(ws, 2, 2, false);

    // Errors
    setCell(ws, 3, 2, makeErrorValue('#DIV/0!'));
    setCell(ws, 4, 2, makeErrorValue('#N/A'));
    setCell(ws, 5, 2, makeErrorValue('#REF!'));
    setCell(ws, 6, 2, makeErrorValue('#VALUE!'));

    // Empty cells (null) — should not be emitted at all
    for (let r = 1; r <= 5; r++) setCell(ws, r, 3, null);

    const result = await writeWorkbook('01-basic-cells.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
