// Tests for getWorksheetAsTextTable — whole-sheet text shortcut.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { getWorksheetAsTextTable } from '../../src/worksheet/text';
import { setCell } from '../../src/worksheet/worksheet';

describe('getWorksheetAsTextTable', () => {
  it('returns the data extent as an ASCII-art table', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'name'); setCell(ws, 1, 2, 'age');
    setCell(ws, 2, 1, 'Alice'); setCell(ws, 2, 2, 30);
    expect(getWorksheetAsTextTable(ws)).toBe(
      ['| name  | age |', '+-------+-----+', '| Alice | 30  |'].join('\n'),
    );
  });

  it('returns "" for an empty worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(getWorksheetAsTextTable(ws)).toBe('');
  });

  it('uses the data extent (sparse layout)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'tl');
    setCell(ws, 3, 3, 'br');
    const txt = getWorksheetAsTextTable(ws);
    // header + sep + 2 data rows = 4 lines
    expect(txt.split('\n').length).toBe(4);
  });
});
