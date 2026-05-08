// Tests for setCellBorderAll / setRangeBorderBox border-preset helpers.

import { describe, expect, it } from 'vitest';
import {
  getCellBorder,
  setCellBorderAll,
  setRangeBorderBox,
} from '../../../src/xlsx/styles/cell-style';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/workbook';
import { setCell } from '../../../src/xlsx/worksheet/worksheet';

describe('setCellBorderAll', () => {
  it('applies the same Side to all four edges', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellBorderAll(wb, c, { style: 'thin' });
    const b = getCellBorder(wb, c);
    expect(b.left?.style).toBe('thin');
    expect(b.right?.style).toBe('thin');
    expect(b.top?.style).toBe('thin');
    expect(b.bottom?.style).toBe('thin');
  });

  it('passes a hex color through to all sides', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellBorderAll(wb, c, { style: 'medium', color: 'FFAA0033' });
    const b = getCellBorder(wb, c);
    expect(b.top?.color?.rgb).toBe('FFAA0033');
    expect(b.bottom?.color?.rgb).toBe('FFAA0033');
  });

  it('default style is "thin"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    setCellBorderAll(wb, c);
    expect(getCellBorder(wb, c).top?.style).toBe('thin');
  });
});

describe('setRangeBorderBox', () => {
  it('only perimeter cells get an edge stroke when no inner is supplied', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeBorderBox(wb, ws, 'B2:D4', { style: 'medium' });
    // Top-left corner: top + left strokes only
    const tl = ws.rows.get(2)?.get(2);
    if (!tl) throw new Error('expected B2');
    const tlBorder = getCellBorder(wb, tl);
    expect(tlBorder.top?.style).toBe('medium');
    expect(tlBorder.left?.style).toBe('medium');
    expect(tlBorder.right?.style).toBeUndefined();
    expect(tlBorder.bottom?.style).toBeUndefined();
    // Top-right corner: top + right
    const tr = ws.rows.get(2)?.get(4);
    if (!tr) throw new Error('expected D2');
    const trBorder = getCellBorder(wb, tr);
    expect(trBorder.top?.style).toBe('medium');
    expect(trBorder.right?.style).toBe('medium');
    expect(trBorder.left?.style).toBeUndefined();
    // Center cell C3 untouched (no inner specified)
    expect(ws.rows.get(3)?.get(3)).toBeUndefined();
  });

  it('with inner: every cell gets perimeter or interior strokes', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeBorderBox(wb, ws, 'A1:C3', { style: 'thick', inner: 'thin' });
    const center = ws.rows.get(2)?.get(2);
    if (!center) throw new Error('expected B2');
    const cb = getCellBorder(wb, center);
    expect(cb.top?.style).toBe('thin');
    expect(cb.bottom?.style).toBe('thin');
    expect(cb.left?.style).toBe('thin');
    expect(cb.right?.style).toBe('thin');
    // Top edge: top stroke is outer (thick), bottom stroke is inner (thin)
    const top = ws.rows.get(1)?.get(2);
    if (!top) throw new Error('expected B1');
    const tb = getCellBorder(wb, top);
    expect(tb.top?.style).toBe('thick');
    expect(tb.bottom?.style).toBe('thin');
  });

  it('1×1 range applies all four outer edges to a single cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeBorderBox(wb, ws, 'A1', { style: 'thick' });
    const c = ws.rows.get(1)?.get(1);
    if (!c) throw new Error('expected A1');
    const b = getCellBorder(wb, c);
    expect(b.top?.style).toBe('thick');
    expect(b.bottom?.style).toBe('thick');
    expect(b.left?.style).toBe('thick');
    expect(b.right?.style).toBe('thick');
  });

  it('color flows to outer + inner uniformly', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeBorderBox(wb, ws, 'A1:B2', { style: 'thin', color: 'FF112233', inner: 'hair' });
    const c = ws.rows.get(1)?.get(1);
    if (!c) throw new Error('expected A1');
    expect(getCellBorder(wb, c).top?.color?.rgb).toBe('FF112233');
  });

  it('preserves existing cell values when a cell is auto-allocated', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 2, 2, 'data');
    setRangeBorderBox(wb, ws, 'A1:C3', { style: 'thin', inner: 'thin' });
    expect(ws.rows.get(2)?.get(2)?.value).toBe('data');
  });
});
