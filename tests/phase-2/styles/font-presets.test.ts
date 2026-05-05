// Tests for font preset cell helpers.

import { describe, expect, it } from 'vitest';
import {
  getCellFont,
  setBold,
  setCellFont,
  setFontColor,
  setFontName,
  setFontSize,
  setItalic,
  setStrikethrough,
  setUnderline,
} from '../../../src/styles/cell-style';
import { makeFont } from '../../../src/styles/fonts';
import { addWorksheet, createWorkbook } from '../../../src/workbook/workbook';
import { setCell } from '../../../src/worksheet/worksheet';

describe('setBold / setItalic / setStrikethrough', () => {
  it('toggle the corresponding boolean flag', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setBold(wb, c);
    setItalic(wb, c);
    setStrikethrough(wb, c);
    const f = getCellFont(wb, c);
    expect(f.bold).toBe(true);
    expect(f.italic).toBe(true);
    expect(f.strike).toBe(true);
  });

  it('passing false flips the flag off', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setBold(wb, c);
    setBold(wb, c, false);
    expect(getCellFont(wb, c).bold).toBe(false);
  });

  it('preserves other font fields', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellFont(wb, c, makeFont({ name: 'Arial', size: 14 }));
    setBold(wb, c);
    const f = getCellFont(wb, c);
    expect(f.name).toBe('Arial');
    expect(f.size).toBe(14);
    expect(f.bold).toBe(true);
  });
});

describe('setUnderline', () => {
  it('default applies "single"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setUnderline(wb, c);
    expect(getCellFont(wb, c).underline).toBe('single');
  });

  it('explicit style passes through', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setUnderline(wb, c, 'double');
    expect(getCellFont(wb, c).underline).toBe('double');
  });

  it('false drops the underline entirely', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setUnderline(wb, c);
    setUnderline(wb, c, false);
    expect(getCellFont(wb, c).underline).toBeUndefined();
  });
});

describe('setFontSize / setFontName / setFontColor', () => {
  it('writes the corresponding field', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setFontSize(wb, c, 18);
    setFontName(wb, c, 'Arial');
    setFontColor(wb, c, 'FF112233');
    const f = getCellFont(wb, c);
    expect(f.size).toBe(18);
    expect(f.name).toBe('Arial');
    expect(f.color?.rgb).toBe('FF112233');
  });

  it('color accepts a Color partial (theme + tint)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setFontColor(wb, c, { theme: 4, tint: 0.4 });
    const f = getCellFont(wb, c);
    expect(f.color?.theme).toBe(4);
    expect(f.color?.tint).toBeCloseTo(0.4);
  });

  it('compose: bold + 18pt Arial red', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setBold(wb, c);
    setFontSize(wb, c, 18);
    setFontName(wb, c, 'Arial');
    setFontColor(wb, c, 'FFFF0000');
    const f = getCellFont(wb, c);
    expect(f.bold).toBe(true);
    expect(f.size).toBe(18);
    expect(f.name).toBe('Arial');
    expect(f.color?.rgb).toBe('FFFF0000');
  });
});
