// Tests for createWorkbookFromCsvBundle — zip of CSVs → multi-sheet Workbook.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  createWorkbookFromCsvBundle,
  getWorkbookAsCsvBundle,
} from '../../src/workbook/workbook';
import { getCellByCoord, setCell } from '../../src/worksheet/worksheet';

const enc = new TextEncoder();

describe('createWorkbookFromCsvBundle', () => {
  it('round-trips through getWorkbookAsCsvBundle', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'a1');
    setCell(a, 1, 2, 'a2');
    setCell(b, 1, 1, 'b1');
    const bundle = getWorkbookAsCsvBundle(wb);
    const wb2 = createWorkbookFromCsvBundle(bundle);
    expect(wb2.sheets.map((s) => s.sheet.title).sort()).toEqual(['A', 'B']);
    const a2 = wb2.sheets.find((s) => s.sheet.title === 'A');
    if (!a2 || a2.kind !== 'worksheet') throw new Error('A missing');
    expect(getCellByCoord(a2.sheet, 'A1')?.value).toBe('a1');
    expect(getCellByCoord(a2.sheet, 'B1')?.value).toBe('a2');
  });

  it('skips non-CSV entries in the zip', () => {
    const bundle = zipSync({
      'Data.csv': enc.encode('a,b\n1,2'),
      'README.txt': enc.encode('not a csv'),
      'image.png': enc.encode('binary'),
    });
    const wb = createWorkbookFromCsvBundle(bundle);
    expect(wb.sheets.map((s) => s.sheet.title)).toEqual(['Data']);
  });

  it('returns an empty workbook for an empty zip bundle', () => {
    const wb = createWorkbookFromCsvBundle(zipSync({}));
    expect(wb.sheets.length).toBe(0);
  });

  it('opts.coerceTypes is forwarded to parseCsvToRange', () => {
    const bundle = zipSync({ 'S.csv': enc.encode('n\n42') });
    const wb = createWorkbookFromCsvBundle(bundle, { coerceTypes: true });
    const ref = wb.sheets[0];
    if (!ref || ref.kind !== 'worksheet') throw new Error('expected a worksheet ref');
    expect(getCellByCoord(ref.sheet, 'A2')?.value).toBe(42);
  });

  it('sanitises Excel-disallowed characters in the source filename', () => {
    // Filename has ":" which Excel disallows in sheet titles.
    const bundle = zipSync({ 'a:b.csv': enc.encode('x') });
    const wb = createWorkbookFromCsvBundle(bundle);
    expect(wb.sheets[0]?.sheet.title).toBe('a_b');
  });
});
