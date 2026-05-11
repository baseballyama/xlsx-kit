// Tests for the approximate autofitColumn / autofitColumns helpers.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  autofitColumn,
  autofitColumns,
  getColumnDimension,
  setCell,
} from '../../src/worksheet/worksheet';

describe('autofitColumn', () => {
  it('sizes a column to max-content-length + default padding (2)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'short');
    setCell(ws, 2, 1, 'a longer string');
    autofitColumn(ws, 1);
    expect(getColumnDimension(ws, 1)?.width).toBe('a longer string'.length + 2);
    expect(getColumnDimension(ws, 1)?.customWidth).toBe(true);
  });

  it('clamps to opts.min when content is shorter', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    autofitColumn(ws, 1, { min: 10 });
    expect(getColumnDimension(ws, 1)?.width).toBe(10);
  });

  it('clamps to opts.max when content is longer', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'x'.repeat(200));
    autofitColumn(ws, 1, { max: 30 });
    expect(getColumnDimension(ws, 1)?.width).toBe(30);
  });

  it('respects custom padding', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, '12345');
    autofitColumn(ws, 1, { padding: 5 });
    expect(getColumnDimension(ws, 1)?.width).toBe(10);
  });

  it('returns undefined and leaves the column untouched when no populated cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'data');
    expect(autofitColumn(ws, 5)).toBeUndefined();
    expect(getColumnDimension(ws, 5)).toBeUndefined();
  });

  it('honours minRow / maxRow window', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a tiny header');
    setCell(ws, 5, 1, 'an extremely long body cell, much longer');
    autofitColumn(ws, 1, { minRow: 1, maxRow: 1 });
    expect(getColumnDimension(ws, 1)?.width).toBe('a tiny header'.length + 2);
  });
});

describe('autofitColumns', () => {
  it('sizes every populated column independently', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'short');
    setCell(ws, 1, 2, 'much longer column 2');
    setCell(ws, 2, 2, 'b');
    autofitColumns(ws);
    expect(getColumnDimension(ws, 1)?.width).toBe('short'.length + 2);
    expect(getColumnDimension(ws, 2)?.width).toBe('much longer column 2'.length + 2);
    expect(getColumnDimension(ws, 3)).toBeUndefined();
  });

  it('clamping flows to per-column widths', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'x');
    setCell(ws, 1, 2, 'x'.repeat(100));
    autofitColumns(ws, { min: 5, max: 25 });
    expect(getColumnDimension(ws, 1)?.width).toBe(5);
    expect(getColumnDimension(ws, 2)?.width).toBe(25);
  });

  it('rich text cells use concatenated run text', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, {
      kind: 'rich-text',
      runs: [{ text: 'Hello ' }, { text: 'world!', font: { b: true } }],
    });
    autofitColumns(ws);
    expect(getColumnDimension(ws, 1)?.width).toBe('Hello world!'.length + 2);
  });
});

describe('autofitColumn font-aware mode', () => {
  it('22pt cell produces ~2× the width of an 11pt cell with the same text', async () => {
    const { setCellFont } = await import('../../src/styles/cell-style');
    const { makeFont } = await import('../../src/styles/fonts');
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'header text');
    setCellFont(wb, c, makeFont({ size: 22 }));
    autofitColumn(ws, 1, { workbook: wb });
    // 'header text'.length === 11 → scaled by 22/11 = 2 → 22 → +2 padding = 24.
    expect(getColumnDimension(ws, 1)?.width).toBe(24);
  });

  it('without workbook, font is ignored (string-length fallback)', async () => {
    const { setCellFont } = await import('../../src/styles/cell-style');
    const { makeFont } = await import('../../src/styles/fonts');
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'header text');
    setCellFont(wb, c, makeFont({ size: 22 }));
    autofitColumn(ws, 1);
    expect(getColumnDimension(ws, 1)?.width).toBe('header text'.length + 2);
  });

  it('autofitColumns with workbook scales each column independently', async () => {
    const { setCellFont } = await import('../../src/styles/cell-style');
    const { makeFont } = await import('../../src/styles/fonts');
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const small = setCell(ws, 1, 1, 'small11');
    const big = setCell(ws, 1, 2, 'big22pt');
    setCellFont(wb, big, makeFont({ size: 22 }));
    // small stays font 0 (default 11pt).
    void small;
    autofitColumns(ws, { workbook: wb });
    expect(getColumnDimension(ws, 1)?.width).toBe('small11'.length + 2);
    // 'big22pt'.length === 7 → ×2 = 14 → +2 = 16.
    expect(getColumnDimension(ws, 2)?.width).toBe(16);
  });
});
