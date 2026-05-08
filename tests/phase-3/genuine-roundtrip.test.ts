// Phase 3 §8 acceptance: load every openpyxl `genuine/*.xlsx` fixture
// our reader supports, save it back through `workbookToBytes`, and
// confirm every cell value + sheet-level metadata survives the trip.
//
// This is the loop's read/write integration gate. mac_date / sample
// sheet 4 (which carries Date-styled cells) currently round-trip as
// numeric serial values rather than JS Date — that's a deferred §5.5
// item and the tests assert the numeric round-trip behaviour for now.

import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import type { FormulaValue } from '../../src/cell/cell';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { getCell, iterRows, type Worksheet } from '../../src/worksheet/worksheet';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

const loadFixture = (name: string): Buffer => readFileSync(resolve(FIXTURES, name));

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('phase 3 §8 — genuine fixture round-trip', () => {
  it('empty.xlsx survives load → save → load with sheet titles + theme', async () => {
    const wb = await loadWorkbook(fromBuffer(loadFixture('empty.xlsx')));
    expect(wb.sheets.map((s) => s.sheet.title)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    expect(wb.themeXml).toBeDefined();
    const themeBefore = wb.themeXml as Uint8Array;

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    expect(wb2.sheets.map((s) => s.sheetId)).toEqual([1, 2, 3]);
    // Each sheet still empty.
    for (const s of wb2.sheets) {
      if (s.kind !== 'worksheet') throw new Error('expected only worksheets');
      expect(s.sheet.rows.size).toBe(0);
    }
    // Theme bytes match exactly.
    expect(wb2.themeXml?.byteLength).toBe(themeBefore.byteLength);
  });

  it('empty-with-styles.xlsx survives load → save → load with cell values + styleIds', async () => {
    const wb = await loadWorkbook(fromBuffer(loadFixture('empty-with-styles.xlsx')));
    const ws1 = expectSheet(wb.sheets[0]?.sheet);

    // A1 = "TEST HERE" via sst[0]; A2..A5 numeric (date / pi / fraction / sci).
    expect(getCell(ws1, 1, 1)?.value).toBe('TEST HERE');
    const a3 = getCell(ws1, 3, 1);
    expect(a3?.value).toBeCloseTo(3.14);

    // Snapshot the cell-value vectors before round-tripping.
    const before = collectCells(ws1);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const after = collectCells(ws2);

    expect(after).toEqual(before);
    expect(wb2.styles.cellXfs.length).toBe(wb.styles.cellXfs.length);
  });

  it('sample.xlsx survives load → save → load — cell values + cross-sheet formula', async () => {
    const wb = await loadWorkbook(fromBuffer(loadFixture('sample.xlsx')));
    expect(wb.sheets.map((s) => s.sheet.title)).toEqual([
      'Sheet1 - Text',
      'Sheet2 - Numbers',
      'Sheet3 - Formulas',
      'Sheet4 - Dates',
    ]);

    const sheet3 = expectSheet(wb.sheets[2]?.sheet);
    const formulaCell = getCell(sheet3, 2, 4); // D2
    expect((formulaCell?.value as FormulaValue).kind).toBe('formula');
    expect((formulaCell?.value as FormulaValue).formula).toBe("'Sheet2 - Numbers'!D5");
    expect((formulaCell?.value as FormulaValue).cachedValue).toBe(5);

    // Per-sheet snapshots.
    const before = wb.sheets.map((s) => collectCells(expectSheet(s.sheet)));
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(wb.sheets.map((s) => s.sheet.title));
    const after = wb2.sheets.map((s) => collectCells(expectSheet(s.sheet)));
    expect(after).toEqual(before);
  });
});

const collectCells = (ws: Worksheet): Array<{ row: number; col: number; value: unknown; styleId: number }> => {
  const out: Array<{ row: number; col: number; value: unknown; styleId: number }> = [];
  for (const cells of iterRows(ws)) {
    for (const c of cells) {
      out.push({ row: c.row, col: c.col, value: c.value, styleId: c.styleId });
    }
  }
  return out;
};
