// Tests for setRangeValues / applyToRange / setRangeStyle.

import { describe, expect, it } from 'vitest';
import {
  getCellAlignment,
  getCellFont,
  makeAlignment,
  makeColor,
  makeFont,
  makePatternFill,
  setRangeStyle,
} from '../../src/xlsx/styles';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import { applyToRange, getCell, setRangeValues } from '../../src/xlsx/worksheet/worksheet';

describe('setRangeValues', () => {
  it('lays values down starting at the top-left of the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeValues(ws, 'B2:D3', [
      ['Name', 'Age', 'City'],
      ['Alice', 30, 'NYC'],
    ]);
    expect(getCell(ws, 2, 2)?.value).toBe('Name');
    expect(getCell(ws, 2, 3)?.value).toBe('Age');
    expect(getCell(ws, 2, 4)?.value).toBe('City');
    expect(getCell(ws, 3, 2)?.value).toBe('Alice');
    expect(getCell(ws, 3, 3)?.value).toBe(30);
    expect(getCell(ws, 3, 4)?.value).toBe('NYC');
  });

  it('skips null/undefined entries', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeValues(ws, 'A1:C2', [
      ['A', null, 'C'],
      [undefined, 'B', undefined],
    ]);
    expect(getCell(ws, 1, 1)?.value).toBe('A');
    expect(getCell(ws, 1, 2)).toBeUndefined();
    expect(getCell(ws, 1, 3)?.value).toBe('C');
    expect(getCell(ws, 2, 1)).toBeUndefined();
    expect(getCell(ws, 2, 2)?.value).toBe('B');
  });
});

describe('applyToRange', () => {
  it('iterates every coord and allocates cells on demand', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    let count = 0;
    applyToRange(ws, 'A1:B3', (cell, r, c) => {
      cell.value = `${r}-${c}`;
      count++;
    });
    expect(count).toBe(6);
    expect(getCell(ws, 1, 1)?.value).toBe('1-1');
    expect(getCell(ws, 3, 2)?.value).toBe('3-2');
  });
});

describe('setRangeStyle', () => {
  it('applies font + alignment to every cell in the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeValues(ws, 'A1:C2', [
      ['A', 'B', 'C'],
      [1, 2, 3],
    ]);
    setRangeStyle(wb, ws, 'A1:C1', {
      font: makeFont({ bold: true, color: makeColor({ rgb: 'FF000080' }) }),
      alignment: makeAlignment({ horizontal: 'center' }),
      fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFCCCCCC' }) }),
    });
    for (const c of [1, 2, 3]) {
      const cell = getCell(ws, 1, c);
      if (!cell) throw new Error('cell missing');
      expect(getCellFont(wb, cell)?.bold).toBe(true);
      expect(getCellAlignment(wb, cell)?.horizontal).toBe('center');
    }
    // Unstyled cells outside the range keep their original styleId
    // (they were created with styleId=0, and setRangeStyle only
    // touched the row 1 cells).
    const unstyled = getCell(ws, 2, 1);
    if (!unstyled) throw new Error('cell missing');
    expect(unstyled.styleId).toBe(0);
  });

  it('allocates blank cells inside the range so the entire box gets the style', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    // No values yet, just style a 3×3 range.
    setRangeStyle(wb, ws, 'A1:C3', {
      fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFFFFF00' }) }),
    });
    expect(getCell(ws, 2, 2)).toBeDefined();
    expect(getCell(ws, 3, 3)).toBeDefined();
  });

  it('empty opts is a no-op', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeStyle(wb, ws, 'A1:B2', {});
    // No cells were allocated.
    expect(getCell(ws, 1, 1)).toBeUndefined();
  });
});