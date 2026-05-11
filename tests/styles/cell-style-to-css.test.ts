// Tests for cellStyleToCss — styled cell → aggregated CSS-property record.

import { describe, expect, it } from 'vitest';
import { makeAlignment } from '../../../src/styles/alignment';
import {
  cellStyleToCss,
  formatAsHeader,
  setBold,
  setCellAlignment,
  setCellBackgroundColor,
  setCellFont,
} from '../../../src/styles/cell-style';
import { makeFont } from '../../../src/styles/fonts';
import { addWorksheet, createWorkbook } from '../../../src/workbook/workbook';
import { setCell } from '../../../src/worksheet/worksheet';

const cellAt = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
  const c = ws.rows.get(row)?.get(col);
  if (!c) throw new Error(`cell ${row},${col} missing`);
  return c;
};

describe('cellStyleToCss', () => {
  it('returns {} for an unstyled cell (styleId === 0, empty stylesheet)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'unstyled');
    expect(cellStyleToCss(wb, cellAt(ws, 1, 1))).toEqual({});
  });

  it('emits font-weight: bold + DEFAULT_FONT props for a setBold-only cell', () => {
    // setBold merges over getCellFont, which falls back to DEFAULT_FONT (Calibri 11)
    // when no font is set. So the CSS includes font-family + font-size from the default.
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'bold');
    setBold(wb, cellAt(ws, 1, 1));
    expect(cellStyleToCss(wb, cellAt(ws, 1, 1))).toEqual({
      'font-family': "'Calibri'",
      'font-size': '11pt',
      'font-weight': 'bold',
    });
  });

  it('merges font + fill for setBold + setCellBackgroundColor', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'mix');
    setBold(wb, cellAt(ws, 1, 1));
    setCellBackgroundColor(wb, cellAt(ws, 1, 1), 'FFFF00');
    const css = cellStyleToCss(wb, cellAt(ws, 1, 1));
    expect(css['font-weight']).toBe('bold');
    expect(css['background-color']).toBe('#FFFF00');
  });

  it('formatAsHeader emits font/fill/border together (no alignment by default)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'Header');
    formatAsHeader(wb, ws, 'A1');
    const css = cellStyleToCss(wb, cellAt(ws, 1, 1));
    expect(css['font-weight']).toBe('bold');
    // formatAsHeader's default fill is 'FF305496' → white-on-dark
    expect(css['background-color']).toBe('#305496');
    expect(css['color']).toBe('#FFFFFF');
    expect(css['border-bottom']).toMatch(/^2px solid /);
  });

  it("collision precedence: alignment overrides font's vertical-align", () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'sup+center');
    setCellFont(wb, cellAt(ws, 1, 1), makeFont({ vertAlign: 'superscript' }));
    setCellAlignment(wb, cellAt(ws, 1, 1), makeAlignment({ vertical: 'center' }));
    expect(cellStyleToCss(wb, cellAt(ws, 1, 1))['vertical-align']).toBe('middle');
  });
});
