// Scenario 28: theme color references with tint modulation, plus the
// indexed-color palette. Output: 28-theme-colors.xlsx
//
// What to verify in Excel:
// - "Theme" tab: 6 fills referencing theme indexes 0..5 (typically
//   bg1/text1/bg2/text2/accent1/accent2). Adjacent rows show the same
//   theme index with tint -0.5 / 0 / +0.4 so the reader can see Excel
//   shading the colour darker / base / lighter.
// - Switch the workbook theme via Page Layout → Themes → Office /
//   Slipstream / Wisp etc. — the cells should re-paint, demonstrating
//   that they are theme-bound rather than RGB-frozen.
// - "Indexed" tab: 8x8 grid of the legacy 64-entry indexed palette
//   (entries 0..63). Useful for spotting any indexed → resolved-RGB
//   bugs. Indices are labelled in column A.

import { describe, expect, it } from 'vitest';
import { addCellXf, addFill, defaultCellXf, makeColor, makePatternFill } from '../../../src/xlsx/styles/index';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { setCell, setColumnWidth } from '../../../src/xlsx/worksheet/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 28 — theme + indexed palette', () => {
  it('writes 28-theme-colors.xlsx', async () => {
    const wb = createWorkbook();

    // ---- Theme tab ----
    const theme = addWorksheet(wb, 'Theme');
    setCell(theme, 1, 1, 'theme idx');
    setCell(theme, 1, 2, 'tint -0.5');
    setCell(theme, 1, 3, 'tint 0');
    setCell(theme, 1, 4, 'tint +0.4');
    setColumnWidth(theme, 1, 12);
    [2, 3, 4].forEach((c) => setColumnWidth(theme, c, 18));

    for (let idx = 0; idx <= 5; idx++) {
      const row = idx + 2;
      setCell(theme, row, 1, `theme ${idx}`);
      const tints = [-0.5, 0, 0.4];
      tints.forEach((tint, i) => {
        const fillId = addFill(
          wb.styles,
          makePatternFill({
            patternType: 'solid',
            fgColor: makeColor({ theme: idx, tint }),
          }),
        );
        const xfId = addCellXf(wb.styles, { ...defaultCellXf(), fillId, applyFill: true });
        const cell = setCell(theme, row, i + 2, '');
        cell.styleId = xfId;
      });
    }

    // ---- Indexed tab ----
    const indexed = addWorksheet(wb, 'Indexed');
    setCell(indexed, 1, 1, 'index');
    setCell(indexed, 1, 2, 'colour');
    for (let i = 0; i < 64; i++) {
      const row = i + 2;
      setCell(indexed, row, 1, i);
      const fillId = addFill(
        wb.styles,
        makePatternFill({
          patternType: 'solid',
          fgColor: makeColor({ indexed: i }),
        }),
      );
      const xfId = addCellXf(wb.styles, { ...defaultCellXf(), fillId, applyFill: true });
      const cell = setCell(indexed, row, 2, '');
      cell.styleId = xfId;
    }

    const result = await writeWorkbook('28-theme-colors.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
