// Tests for alignment-preset cell helpers.

import { describe, expect, it } from 'vitest';
import {
  alignCellHorizontal,
  alignCellVertical,
  centerCell,
  getCellAlignment,
  indentCell,
  rotateCellText,
  setCellAlignment,
  wrapCellText,
} from '../../../src/styles/cell-style';
import { makeAlignment } from '../../../src/styles/alignment';
import { addWorksheet, createWorkbook } from '../../../src/workbook/workbook';
import { setCell } from '../../../src/worksheet/worksheet';

describe('centerCell', () => {
  it('sets horizontal + vertical to "center"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    centerCell(wb, c);
    const a = getCellAlignment(wb, c);
    expect(a.horizontal).toBe('center');
    expect(a.vertical).toBe('center');
  });

  it('preserves other alignment fields', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellAlignment(wb, c, makeAlignment({ wrapText: true, indent: 2 }));
    centerCell(wb, c);
    const a = getCellAlignment(wb, c);
    expect(a.horizontal).toBe('center');
    expect(a.wrapText).toBe(true);
    expect(a.indent).toBe(2);
  });
});

describe('wrapCellText', () => {
  it('toggles wrapText true / false', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    wrapCellText(wb, c);
    expect(getCellAlignment(wb, c).wrapText).toBe(true);
    wrapCellText(wb, c, false);
    expect(getCellAlignment(wb, c).wrapText).toBe(false);
  });
});

describe('alignCellHorizontal / alignCellVertical', () => {
  it('set the corresponding axis without disturbing the other', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    alignCellHorizontal(wb, c, 'right');
    alignCellVertical(wb, c, 'top');
    const a = getCellAlignment(wb, c);
    expect(a.horizontal).toBe('right');
    expect(a.vertical).toBe('top');
  });
});

describe('rotateCellText', () => {
  it('accepts 0..180 and the special 255', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    rotateCellText(wb, c, 90);
    expect(getCellAlignment(wb, c).textRotation).toBe(90);
    rotateCellText(wb, c, 255);
    expect(getCellAlignment(wb, c).textRotation).toBe(255);
  });

  it('rejects out-of-range values', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    expect(() => rotateCellText(wb, c, 200)).toThrow(/0\.\.180 or 255/);
    expect(() => rotateCellText(wb, c, -10)).toThrow(/0\.\.180 or 255/);
  });
});

describe('indentCell', () => {
  it('sets the indent level', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    indentCell(wb, c, 3);
    expect(getCellAlignment(wb, c).indent).toBe(3);
  });
});
