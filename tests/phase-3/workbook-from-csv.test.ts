// Tests for createWorkbookFromCsv — one-shot CSV → Workbook constructor.

import { describe, expect, it } from 'vitest';
import { createWorkbookFromCsv } from '../../src/workbook/workbook';
import { getCellByCoord } from '../../src/worksheet/worksheet';

describe('createWorkbookFromCsv', () => {
  it('creates a 1-sheet workbook with the parsed CSV in Sheet1', () => {
    const wb = createWorkbookFromCsv('a,b\n1,2');
    expect(wb.sheets.length).toBe(1);
    expect(wb.sheets[0]?.sheet.title).toBe('Sheet1');
    const ref = wb.sheets[0];
    if (!ref || ref.kind !== 'worksheet') throw new Error('expected a worksheet ref');
    expect(getCellByCoord(ref.sheet, 'A1')?.value).toBe('a');
    expect(getCellByCoord(ref.sheet, 'B2')?.value).toBe('2');
  });

  it('honours opts.sheetTitle', () => {
    const wb = createWorkbookFromCsv('x', { sheetTitle: 'Imported' });
    expect(wb.sheets[0]?.sheet.title).toBe('Imported');
  });

  it('opts.coerceTypes forwards to parseCsvToRange', () => {
    const wb = createWorkbookFromCsv('a,b,c\n1,true,foo', { coerceTypes: true });
    const ref = wb.sheets[0];
    if (!ref || ref.kind !== 'worksheet') throw new Error('expected a worksheet ref');
    expect(getCellByCoord(ref.sheet, 'A2')?.value).toBe(1);
    expect(getCellByCoord(ref.sheet, 'B2')?.value).toBe(true);
    expect(getCellByCoord(ref.sheet, 'C2')?.value).toBe('foo');
  });

  it('returns a workbook with an empty sheet for empty CSV input', () => {
    const wb = createWorkbookFromCsv('');
    expect(wb.sheets.length).toBe(1);
    const ref = wb.sheets[0];
    if (!ref || ref.kind !== 'worksheet') throw new Error('expected a worksheet ref');
    expect(ref.sheet.rows.size).toBe(0);
  });
});
