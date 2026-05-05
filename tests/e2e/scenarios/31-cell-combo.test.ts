// Scenario 31: a single focal cell carrying every cell-axis decoration
// at once — merge + style (font + fill + border + alignment) + comment +
// hyperlink + data validation. Output: 31-cell-combo.xlsx
//
// What to verify in Excel:
// - B2:D4 is merged into one big block. The block is the focal cell.
// - The merged cell shows "Click me ▶" centered, in bold red Calibri
//   16pt over a yellow fill bordered with thick black on all sides.
// - Hovering shows a comment "All-in-one demo cell" by author "QA".
// - Clicking the cell opens https://example.com/ in the browser.
// - The cell also has a data-validation list (Yes / No / Maybe) — the
//   dropdown arrow appears on selection. The validation is informational
//   ("warning" style) so even a non-matching paste isn't rejected.
// - Cells outside the merge are normal.

import { describe, expect, it } from 'vitest';
import {
  addBorder,
  addDataValidation,
  addFill,
  addFont,
  addWorksheet,
  addCellXf,
  createWorkbook,
  defaultCellXf,
  makeAlignment,
  makeBorder,
  makeColor,
  makeFont,
  makePatternFill,
  makeSide,
  mergeCells,
  setCell,
  setColumnWidth,
  setComment,
  setHyperlink,
  setRowHeight,
} from '../../../src/index';
import { parseMultiCellRange } from '../../../src/worksheet/cell-range';
import { writeWorkbook } from '../_helpers';

describe('e2e 31 — single cell, every decoration at once', () => {
  it('writes 31-cell-combo.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Combo');

    // Build a fancy style: bold red Calibri 16pt + yellow fill + thick
    // black border + center align.
    const fontId = addFont(
      wb.styles,
      makeFont({
        name: 'Calibri',
        size: 16,
        bold: true,
        color: makeColor({ rgb: 'FFC00000' }),
      }),
    );
    const fillId = addFill(
      wb.styles,
      makePatternFill({
        patternType: 'solid',
        fgColor: makeColor({ rgb: 'FFFFFF00' }),
      }),
    );
    const thick = makeSide({ style: 'thick', color: makeColor({ rgb: 'FF000000' }) });
    const borderId = addBorder(wb.styles, makeBorder({ left: thick, right: thick, top: thick, bottom: thick }));
    const xfId = addCellXf(wb.styles, {
      ...defaultCellXf(),
      fontId,
      fillId,
      borderId,
      applyFont: true,
      applyFill: true,
      applyBorder: true,
      alignment: makeAlignment({ horizontal: 'center', vertical: 'center', wrapText: true }),
      applyAlignment: true,
    });

    // Merge B2:D4 and place the focal value in B2.
    mergeCells(ws, 'B2:D4');
    const focal = setCell(ws, 2, 2, 'Click me ▶');
    focal.styleId = xfId;

    // Make the merged block visibly large.
    [2, 3, 4].forEach((c) => setColumnWidth(ws, c, 18));
    [2, 3, 4].forEach((r) => setRowHeight(ws, r, 30));

    setHyperlink(ws, 'B2', {
      target: 'https://example.com/',
      tooltip: 'All-in-one demo cell',
      display: 'Click me ▶',
    });

    setComment(ws, { ref: 'B2', author: 'QA', text: 'All-in-one demo cell' });

    addDataValidation(ws, {
      sqref: parseMultiCellRange('B2:D4'),
      type: 'list',
      operator: 'between',
      errorStyle: 'warning',
      showErrorMessage: true,
      showInputMessage: true,
      promptTitle: 'Pick one',
      prompt: 'Yes / No / Maybe',
      formula1: '"Yes,No,Maybe"',
      allowBlank: true,
    });

    setCell(ws, 6, 2, 'Other cells are normal — only B2:D4 is loaded up.');

    const result = await writeWorkbook('31-cell-combo.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
