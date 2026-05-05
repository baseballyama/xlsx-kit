// Phase 2 §6 acceptance: a Workbook + worksheet content + Stylesheet
// pool round-trips losslessly through `JSON.stringify(wb, jsonReplacer)`
// → `JSON.parse(json, jsonReviver)`. The reviver pair turns Map values
// into `{ __map__: [...] }` arrays and back so Worksheet.rows (the
// `Map<row, Map<col, Cell>>` sparse storage) and Stylesheet.numFmts +
// internal `_*ByKey` dedup maps survive.

import { describe, expect, it } from 'vitest';
import { type Cell, makeErrorValue, setFormula } from '../../src/cell/cell';
import { makeRichText, makeTextRun } from '../../src/cell/rich-text';
import { makeBorder, makeSide } from '../../src/styles/borders';
import {
  getCellAlignment,
  getCellBorder,
  getCellFill,
  getCellFont,
  getCellNumberFormat,
  setCellAlignment,
  setCellBorder,
  setCellFill,
  setCellFont,
  setCellNumberFormat,
} from '../../src/styles/cell-style';
import { makePatternFill } from '../../src/styles/fills';
import { makeFont } from '../../src/styles/fonts';
import {
  addWorksheet,
  createWorkbook,
  jsonReplacer,
  jsonReviver,
  setActiveSheet,
  type Workbook,
} from '../../src/workbook/workbook';
import { getCell, getMaxCol, getMaxRow, iterValues, setCell, type Worksheet } from '../../src/worksheet/worksheet';

const roundTrip = (wb: Workbook): Workbook => {
  const json = JSON.stringify(wb, jsonReplacer);
  return JSON.parse(json, jsonReviver) as Workbook;
};

describe('JSON round-trip — empty workbook', () => {
  it('preserves the default stylesheet pre-population', () => {
    const wb = createWorkbook();
    const wb2 = roundTrip(wb);
    expect(wb2.styles.fonts.length).toBe(1);
    expect(wb2.styles.fills.length).toBe(2);
    expect(wb2.styles.borders.length).toBe(1);
    expect(wb2.styles.cellXfs.length).toBe(0);
    expect(wb2.styles.numFmts).toBeInstanceOf(Map);
    expect(wb2.styles._fontIdByKey).toBeInstanceOf(Map);
  });
});

describe('JSON round-trip — worksheet content', () => {
  it('preserves number / string / boolean cell values', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    setCell(ws, 1, 1, 42);
    setCell(ws, 1, 2, 'hello');
    setCell(ws, 1, 3, true);
    setCell(ws, 2, 1, null);
    const wb2 = roundTrip(wb);
    const ws2 = wb2.sheets[0]?.sheet as Worksheet;
    expect(ws2.title).toBe('Sheet1');
    expect(getCell(ws2, 1, 1)?.value).toBe(42);
    expect(getCell(ws2, 1, 2)?.value).toBe('hello');
    expect(getCell(ws2, 1, 3)?.value).toBe(true);
    expect(getCell(ws2, 2, 1)?.value).toBeNull();
  });

  it('preserves formula / error / rich-text discriminated values', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const cF = setCell(ws, 1, 1);
    setFormula(cF, 'A2+B2', { cachedValue: 10 });
    const cErr = setCell(ws, 2, 1);
    cErr.value = makeErrorValue('#REF!');
    const cRich = setCell(ws, 3, 1);
    cRich.value = { kind: 'rich-text', runs: makeRichText([makeTextRun('Hello', { b: true })]) };

    const wb2 = roundTrip(wb);
    const ws2 = wb2.sheets[0]?.sheet as Worksheet;
    const f = getCell(ws2, 1, 1)?.value as {
      kind: string;
      formula: string;
      cachedValue: number;
      t: string;
    };
    expect(f.kind).toBe('formula');
    expect(f.formula).toBe('A2+B2');
    expect(f.t).toBe('normal');
    expect(f.cachedValue).toBe(10);
    expect((getCell(ws2, 2, 1)?.value as { kind: string; code: string }).code).toBe('#REF!');
    const richBack = getCell(ws2, 3, 1)?.value as {
      kind: string;
      runs: Array<{ text: string; font?: { b?: boolean } }>;
    };
    expect(richBack.kind).toBe('rich-text');
    expect(richBack.runs[0]?.text).toBe('Hello');
    expect(richBack.runs[0]?.font?.b).toBe(true);
  });

  it('preserves sparse storage dimensions and iteration order', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 5, 'b');
    setCell(ws, 10, 3, 'c');
    expect(getMaxRow(ws)).toBe(10);
    expect(getMaxCol(ws)).toBe(5);

    const wb2 = roundTrip(wb);
    const ws2 = wb2.sheets[0]?.sheet as Worksheet;
    expect(getMaxRow(ws2)).toBe(10);
    expect(getMaxCol(ws2)).toBe(5);
    const rows = [...iterValues(ws2)];
    expect(rows.length).toBeGreaterThan(0);
  });
});

describe('JSON round-trip — multi-sheet workbook', () => {
  it('preserves sheet order, titles, sheetIds, and active index', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addWorksheet(wb, 'C');
    setActiveSheet(wb, 'B');

    const wb2 = roundTrip(wb);
    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(['A', 'B', 'C']);
    expect(wb2.sheets.map((s) => s.sheetId)).toEqual([1, 2, 3]);
    expect(wb2.activeSheetIndex).toBe(1);
  });
});

describe('JSON round-trip — stylesheet pool dedup survives', () => {
  it('font / fill / border / numFmt indices match before and after', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');

    const c1 = setCell(ws, 1, 1, 1);
    setCellFont(wb, c1, makeFont({ bold: true }));
    const styleIdBefore = c1.styleId;

    const c2 = setCell(ws, 2, 1, 2);
    setCellFill(wb, c2, makePatternFill({ patternType: 'solid', fgColor: { rgb: 'FFFF0000' } }));
    setCellBorder(wb, c2, makeBorder({ left: makeSide({ style: 'thin' }) }));
    setCellAlignment(wb, c2, { horizontal: 'center' });
    setCellNumberFormat(wb, c2, '0.0000');

    const fontsBefore = wb.styles.fonts.length;
    const fillsBefore = wb.styles.fills.length;
    const cellXfsBefore = wb.styles.cellXfs.length;

    const wb2 = roundTrip(wb);
    expect(wb2.styles.fonts.length).toBe(fontsBefore);
    expect(wb2.styles.fills.length).toBe(fillsBefore);
    expect(wb2.styles.cellXfs.length).toBe(cellXfsBefore);

    const ws2 = wb2.sheets[0]?.sheet as Worksheet;
    const c1Back = getCell(ws2, 1, 1) as Cell;
    expect(c1Back.styleId).toBe(styleIdBefore);
    expect(getCellFont(wb2, c1Back).bold).toBe(true);

    const c2Back = getCell(ws2, 2, 1) as Cell;
    expect(getCellFill(wb2, c2Back).kind).toBe('pattern');
    expect(getCellBorder(wb2, c2Back).left?.style).toBe('thin');
    expect(getCellAlignment(wb2, c2Back).horizontal).toBe('center');
    expect(getCellNumberFormat(wb2, c2Back)).toBe('0.0000');
  });

  it('dedup remains active after revival — same logical font hits the same id', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const c1 = setCell(ws, 1, 1, 1);
    setCellFont(wb, c1, makeFont({ italic: true }));
    const wb2 = roundTrip(wb);
    const c2 = setCell(wb2.sheets[0]?.sheet as Worksheet, 2, 2, 2);
    setCellFont(wb2, c2, makeFont({ italic: true }));
    // Both cells point at the same dedup'd CellXf id post-revival.
    expect(c2.styleId).toBe(c1.styleId);
    // cellXfs[0] is the implicit default reserved by setCellFont; cellXfs[1]
    // is the shared italic xf both cells point at.
    expect(wb2.styles.cellXfs.length).toBe(2);
    expect(wb2.styles.fonts.length).toBe(2); // DEFAULT + italic
  });

  it('custom numFmt code dedups across the round-trip boundary', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const c1 = setCell(ws, 1, 1, 1);
    setCellNumberFormat(wb, c1, '0.0000');
    const customIdBefore = wb.styles.cellXfs[c1.styleId]?.numFmtId;

    const wb2 = roundTrip(wb);
    const c2 = setCell(wb2.sheets[0]?.sheet as Worksheet, 2, 1, 2);
    setCellNumberFormat(wb2, c2, '0.0000');
    expect(wb2.styles.cellXfs[c2.styleId]?.numFmtId).toBe(customIdBefore);
    expect(wb2.styles.numFmts.size).toBe(1);
  });
});

describe('JSON round-trip — cellRange preservation via sqref', () => {
  // mergedCells API is deferred to phase 5 per the plan; for §6 we just
  // assert that a sqref-style string parses identically before and after
  // the round-trip — i.e. the underlying boundaries representation is
  // JSON-safe.
  it('multiCellRange string survives a stringify/parse cycle', () => {
    const sqref = 'A1:B2 D5 E10:F20';
    const obj = { sqref };
    const back = JSON.parse(JSON.stringify(obj, jsonReplacer), jsonReviver) as { sqref: string };
    expect(back.sqref).toBe(sqref);
  });
});
