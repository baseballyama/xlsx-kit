// Tests for setCellRichText.

import { describe, expect, it } from 'vitest';
import { isRichTextValue } from '../../src/xlsx/cell/cell';
import { makeRichText } from '../../src/xlsx/cell/rich-text';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { setCellRichText } from '../../src/xlsx/worksheet/worksheet';

describe('setCellRichText', () => {
  it('writes a rich-text value built from inline run objects', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCellRichText(ws, 1, 1, [
      { text: 'Hello ' },
      { text: 'world', font: { b: true } },
    ]);
    expect(isRichTextValue(c.value)).toBe(true);
    if (isRichTextValue(c.value)) {
      expect(c.value.runs.length).toBe(2);
      expect(c.value.runs[0]?.text).toBe('Hello ');
      expect(c.value.runs[1]?.font?.b).toBe(true);
    }
  });

  it('accepts a pre-built RichText (frozen TextRun array)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const rt = makeRichText([{ text: 'hi' }]);
    const c = setCellRichText(ws, 1, 1, rt);
    expect(isRichTextValue(c.value)).toBe(true);
  });

  it('honours optional styleId', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCellRichText(ws, 1, 1, [{ text: 'x' }], 7);
    expect(c.styleId).toBe(7);
  });
});
