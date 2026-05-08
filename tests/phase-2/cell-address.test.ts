// Tests for getCellAddress / formatSheetQualifiedRef — sheet-qualified A1.

import { describe, expect, it } from 'vitest';
import { formatSheetQualifiedRef, parseSheetRange } from '../../src/utils/coordinate';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { getCellAddress, setCell } from '../../src/xlsx/worksheet/worksheet';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('formatSheetQualifiedRef', () => {
  it('emits bare title for plain alphanumeric_underscore names', () => {
    expect(formatSheetQualifiedRef('Sheet1', 'A1')).toBe('Sheet1!A1');
    expect(formatSheetQualifiedRef('Data_2024', 'B5')).toBe('Data_2024!B5');
  });

  it('quotes titles containing spaces', () => {
    expect(formatSheetQualifiedRef('Quarter 1', 'A1')).toBe("'Quarter 1'!A1");
  });

  it("doubles internal apostrophes inside quoted titles", () => {
    expect(formatSheetQualifiedRef("Bob's Sheet", 'A1')).toBe("'Bob''s Sheet'!A1");
  });

  it('quotes titles starting with a digit (cannot be a bare ref)', () => {
    expect(formatSheetQualifiedRef('2024Data', 'A1')).toBe("'2024Data'!A1");
  });

  it('quotes titles containing punctuation (-, ., comma)', () => {
    expect(formatSheetQualifiedRef('Q1-Sales', 'A1')).toBe("'Q1-Sales'!A1");
    expect(formatSheetQualifiedRef('Sheet.1', 'A1')).toBe("'Sheet.1'!A1");
  });

  it('round-trips through parseSheetRange', () => {
    for (const title of ['Sheet1', 'Quarter 1', "Bob's Sheet", '2024', 'Q1-Sales']) {
      const formatted = formatSheetQualifiedRef(title, 'B5');
      const parsed = parseSheetRange(formatted);
      expect(parsed.sheet).toBe(title);
      expect(parsed.range).toBe('B5');
    }
  });
});

describe('getCellAddress', () => {
  it('returns the sheet-qualified address for a cell on a plain-named sheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    setCell(ws, 1, 1, 'a');
    expect(getCellAddress(ws, cellAt(ws, 1, 1))).toBe('Data!A1');
  });

  it('quotes the sheet name when it contains spaces', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'My Sheet');
    setCell(ws, 5, 7, 'x');
    expect(getCellAddress(ws, cellAt(ws, 5, 7))).toBe("'My Sheet'!G5");
  });
});
