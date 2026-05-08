// Phase 2 acceptance gate per docs/plan/04-core-model.md §8.
// Sample script: createWorkbook → 100 cells written → JSON-serialised
// → JSON-revived → identity check across every cell + the styling pool.
// This complements the targeted json-roundtrip.test.ts by exercising
// scale + a styling mix.

import { describe, expect, it } from 'vitest';
import { setCellFont } from '../../src/styles/cell-style';
import { makeFont } from '../../src/styles/fonts';
import { addWorksheet, createWorkbook, jsonReplacer, jsonReviver, type Workbook } from '../../src/workbook/workbook';
import { getCell, setCell, type Worksheet } from '../../src/worksheet/worksheet';

describe('phase 2 §8 — 100 cell write / JSON round-trip / identity check', () => {
  it('round-trips losslessly with mixed values + styles', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'AcceptSheet');

    // 10 × 10 grid, alternating value kinds + a couple of styles.
    const boldFont = makeFont({ bold: true });
    const italicFont = makeFont({ italic: true });
    for (let row = 1; row <= 10; row++) {
      for (let col = 1; col <= 10; col++) {
        const idx = (row - 1) * 10 + col;
        // Value kinds rotate so we cover number / string / bool / null.
        let value: number | string | boolean | null;
        const m = idx % 4;
        if (m === 0) value = idx;
        else if (m === 1) value = `cell-${idx}`;
        else if (m === 2) value = idx % 2 === 0;
        else value = null;

        const c = setCell(ws, row, col, value);
        if (col === 1) setCellFont(wb, c, boldFont);
        else if (col === 10) setCellFont(wb, c, italicFont);
      }
    }

    expect([...ws.rows.keys()].length).toBe(10);

    const json = JSON.stringify(wb, jsonReplacer);
    const wb2 = JSON.parse(json, jsonReviver) as Workbook;
    const ws2 = wb2.sheets[0]?.sheet as Worksheet;

    // Stylesheet pool sizes match exactly.
    expect(wb2.styles.fonts.length).toBe(wb.styles.fonts.length);
    expect(wb2.styles.fills.length).toBe(wb.styles.fills.length);
    expect(wb2.styles.borders.length).toBe(wb.styles.borders.length);
    expect(wb2.styles.cellXfs.length).toBe(wb.styles.cellXfs.length);

    // Every cell value + styleId matches.
    for (let row = 1; row <= 10; row++) {
      for (let col = 1; col <= 10; col++) {
        const before = getCell(ws, row, col);
        const after = getCell(ws2, row, col);
        expect(after).toBeDefined();
        expect(after?.row).toBe(before?.row);
        expect(after?.col).toBe(before?.col);
        expect(after?.value).toEqual(before?.value);
        expect(after?.styleId).toBe(before?.styleId);
      }
    }
  });
});
