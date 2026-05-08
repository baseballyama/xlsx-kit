// Scenario 05: Font / Fill / Border / Alignment / NumberFormat /
// Protection — every style axis on a single cell, plus a grid that
// shows each axis independently.
//
// Output: 05-styles.xlsx
//
// What to verify in Excel:
// - Sheet "AllAxes" A1 should be bold red Calibri 14pt on yellow fill,
//   with all four borders thick black, centre-aligned, format "#,##0.00",
//   protection.locked=false (visible only in protected sheets).
// - Sheet "Fonts" A1..A6: regular / bold / italic / underline / strike /
//   bold+italic+underline.
// - Sheet "Fills" A1..A4: no fill / yellow solid / pattern / gradient.
// - Sheet "Borders" A1..A4: thin / medium / thick / double on all sides.
// - Sheet "Align" A1..A4: left / center / right / justify; B1..B4: top /
//   middle / bottom; wrap text in C1.
// - Sheet "NumFmt" A1..A6: #,##0 / #,##0.00 / 0% / "$"#,##0.00 /
//   yyyy-mm-dd / [h]:mm:ss applied to a sample value.

import { describe, expect, it } from 'vitest';
import { addBorder, addCellXf, addFill, addFont, addNumFmt, defaultCellXf, makeAlignment, makeBorder, makeColor, makeFont, makePatternFill, makeProtection, makeSide } from '../../../src/styles/index';
import { addWorksheet, createWorkbook } from '../../../src/workbook/index';
import { setCell } from '../../../src/worksheet/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 05 — styles', () => {
  it('writes 05-styles.xlsx', async () => {
    const wb = createWorkbook();
    addCellXf(wb.styles, defaultCellXf()); // reserve slot 0 for default

    // -------- AllAxes -----------------------------------------------------
    const all = addWorksheet(wb, 'AllAxes');
    const fontRedBold = addFont(wb.styles, makeFont({ name: 'Calibri', size: 14, bold: true, color: makeColor({ rgb: 'FFFF0000' }) }));
    const fillYellow = addFill(wb.styles, makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFFFFF00' }) }));
    const blackThick = makeSide({ style: 'thick', color: makeColor({ rgb: 'FF000000' }) });
    const borderBox = addBorder(wb.styles, makeBorder({ left: blackThick, right: blackThick, top: blackThick, bottom: blackThick }));
    const numFmt = addNumFmt(wb.styles, '#,##0.00');
    const xf = addCellXf(wb.styles, {
      ...defaultCellXf(),
      fontId: fontRedBold,
      fillId: fillYellow,
      borderId: borderBox,
      numFmtId: numFmt,
      alignment: makeAlignment({ horizontal: 'center', vertical: 'center' }),
      protection: makeProtection({ locked: false }),
      applyFont: true,
      applyFill: true,
      applyBorder: true,
      applyNumberFormat: true,
      applyAlignment: true,
      applyProtection: true,
    });
    setCell(all, 2, 2, 1234.5678, xf);

    // -------- Fonts ------------------------------------------------------
    const fonts = addWorksheet(wb, 'Fonts');
    const fontVariants = [
      { label: 'regular', font: makeFont({ name: 'Calibri', size: 11 }) },
      { label: 'bold', font: makeFont({ name: 'Calibri', size: 11, bold: true }) },
      { label: 'italic', font: makeFont({ name: 'Calibri', size: 11, italic: true }) },
      { label: 'underline', font: makeFont({ name: 'Calibri', size: 11, underline: 'single' }) },
      { label: 'strike', font: makeFont({ name: 'Calibri', size: 11, strike: true }) },
      { label: 'bold+italic+underline', font: makeFont({ name: 'Calibri', size: 11, bold: true, italic: true, underline: 'single' }) },
    ];
    for (let i = 0; i < fontVariants.length; i++) {
      const v = fontVariants[i];
      if (!v) continue;
      const r = i + 1;
      setCell(fonts, r, 1, v.label);
      const fid = addFont(wb.styles, v.font);
      const xfid = addCellXf(wb.styles, { ...defaultCellXf(), fontId: fid, applyFont: true });
      setCell(fonts, r, 2, 'AaBbCc 123', xfid);
    }

    // -------- Fills ------------------------------------------------------
    const fills = addWorksheet(wb, 'Fills');
    const fillVariants = [
      { label: 'no fill', fill: makePatternFill({ patternType: 'none' }) },
      { label: 'yellow solid', fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFFFFF00' }) }) },
      { label: 'darkVertical pattern', fill: makePatternFill({ patternType: 'darkVertical', fgColor: makeColor({ rgb: 'FF0070C0' }), bgColor: makeColor({ rgb: 'FFFFFF00' }) }) },
    ];
    for (let i = 0; i < fillVariants.length; i++) {
      const v = fillVariants[i];
      if (!v) continue;
      const r = i + 1;
      setCell(fills, r, 1, v.label);
      const fid = addFill(wb.styles, v.fill);
      const xfid = addCellXf(wb.styles, { ...defaultCellXf(), fillId: fid, applyFill: true });
      setCell(fills, r, 2, 'sample', xfid);
    }

    // -------- Borders ----------------------------------------------------
    const borders = addWorksheet(wb, 'Borders');
    const borderStyles: Array<'thin' | 'medium' | 'thick' | 'double'> = ['thin', 'medium', 'thick', 'double'];
    for (let i = 0; i < borderStyles.length; i++) {
      const s = borderStyles[i];
      if (!s) continue;
      const side = makeSide({ style: s, color: makeColor({ rgb: 'FF000000' }) });
      const bid = addBorder(wb.styles, makeBorder({ left: side, right: side, top: side, bottom: side }));
      const xfid = addCellXf(wb.styles, { ...defaultCellXf(), borderId: bid, applyBorder: true });
      const r = i + 1;
      setCell(borders, r, 1, s);
      setCell(borders, r, 2, 'B', xfid);
    }

    // -------- Alignment --------------------------------------------------
    const align = addWorksheet(wb, 'Align');
    const horizontals: Array<'left' | 'center' | 'right' | 'justify'> = ['left', 'center', 'right', 'justify'];
    horizontals.forEach((h, i) => {
      const xfid = addCellXf(wb.styles, { ...defaultCellXf(), alignment: makeAlignment({ horizontal: h }), applyAlignment: true });
      setCell(align, i + 1, 1, h);
      setCell(align, i + 1, 2, 'aligned text content', xfid);
    });
    const verticals: Array<'top' | 'center' | 'bottom'> = ['top', 'center', 'bottom'];
    verticals.forEach((v, i) => {
      const xfid = addCellXf(wb.styles, { ...defaultCellXf(), alignment: makeAlignment({ vertical: v }), applyAlignment: true });
      setCell(align, 1, 3 + i, v);
      setCell(align, 2, 3 + i, 'V', xfid);
    });
    const wrapXf = addCellXf(wb.styles, { ...defaultCellXf(), alignment: makeAlignment({ wrapText: true }), applyAlignment: true });
    setCell(align, 1, 6, 'wrap');
    setCell(align, 2, 6, 'a quite long text that should wrap when the column is narrow', wrapXf);

    // -------- NumFmt -----------------------------------------------------
    const num = addWorksheet(wb, 'NumFmt');
    const numFmts = ['#,##0', '#,##0.00', '0%', '"$"#,##0.00', 'yyyy-mm-dd', '[h]:mm:ss'];
    numFmts.forEach((code, i) => {
      const id = addNumFmt(wb.styles, code);
      const xfid = addCellXf(wb.styles, { ...defaultCellXf(), numFmtId: id, applyNumberFormat: true });
      setCell(num, i + 1, 1, code);
      setCell(num, i + 1, 2, code === '0%' ? 0.5 : code.startsWith('y') ? 45292 : code === '[h]:mm:ss' ? 1.5 : 1234.5678, xfid);
    });

    const result = await writeWorkbook('05-styles.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
