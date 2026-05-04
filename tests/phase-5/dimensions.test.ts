import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  getColumnDimension,
  getRowDimension,
  hideColumn,
  hideRow,
  setCell,
  setColumnDimension,
  setColumnWidth,
  setRowDimension,
  setRowHeight,
  type Worksheet,
} from '../../src/worksheet/worksheet';

const expectSheet = (ws: Worksheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  return ws;
};

describe('column dimension API', () => {
  it('setColumnWidth populates width + customWidth', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    setColumnWidth(ws, 3, 18.5);
    const dim = getColumnDimension(ws, 3);
    expect(dim?.width).toBe(18.5);
    expect(dim?.customWidth).toBe(true);
    expect(dim?.min).toBe(3);
    expect(dim?.max).toBe(3);
  });

  it('setColumnDimension overwrites the prior entry covering that column', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    setColumnWidth(ws, 5, 12);
    setColumnDimension(ws, 5, { width: 20, hidden: true });
    expect(getColumnDimension(ws, 5)?.width).toBe(20);
    expect(getColumnDimension(ws, 5)?.hidden).toBe(true);
    // No leftover entries.
    expect(ws.columnDimensions.size).toBe(1);
  });

  it('hideColumn flips hidden=true', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    hideColumn(ws, 2);
    expect(getColumnDimension(ws, 2)?.hidden).toBe(true);
  });
});

describe('row dimension API', () => {
  it('setRowHeight populates height + customHeight', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'R');
    setRowHeight(ws, 7, 24);
    const dim = getRowDimension(ws, 7);
    expect(dim?.height).toBe(24);
    expect(dim?.customHeight).toBe(true);
  });

  it('hideRow flips hidden=true', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'R');
    hideRow(ws, 4);
    expect(getRowDimension(ws, 4)?.hidden).toBe(true);
  });

  it('setRowDimension is the lower-level form', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'R');
    setRowDimension(ws, 1, { height: 30, hidden: true, outlineLevel: 2 });
    const dim = getRowDimension(ws, 1);
    expect(dim?.height).toBe(30);
    expect(dim?.hidden).toBe(true);
    expect(dim?.outlineLevel).toBe(2);
  });
});

describe('column / row dimensions round-trip through saveWorkbook → loadWorkbook', () => {
  it('preserves width / hidden columns', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    setColumnWidth(ws, 1, 12.5);
    hideColumn(ws, 4);
    setColumnDimension(ws, 7, { width: 5, bestFit: true });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(getColumnDimension(ws2, 1)?.width).toBe(12.5);
    expect(getColumnDimension(ws2, 1)?.customWidth).toBe(true);
    expect(getColumnDimension(ws2, 4)?.hidden).toBe(true);
    expect(getColumnDimension(ws2, 7)?.width).toBe(5);
    expect(getColumnDimension(ws2, 7)?.bestFit).toBe(true);
  });

  it('preserves row heights + hidden rows + dimension-only rows', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'R');
    setCell(ws, 1, 1, 'top');
    setRowHeight(ws, 1, 28);
    hideRow(ws, 5);
    // Row 10 has only a height entry (no cells) — should still emit + load.
    setRowHeight(ws, 10, 50);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(getRowDimension(ws2, 1)?.height).toBe(28);
    expect(getRowDimension(ws2, 5)?.hidden).toBe(true);
    expect(getRowDimension(ws2, 10)?.height).toBe(50);
  });

  it('preserves defaultColumnWidth / defaultRowHeight', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'D');
    ws.defaultColumnWidth = 10;
    ws.defaultRowHeight = 18;
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.defaultColumnWidth).toBe(10);
    expect(ws2.defaultRowHeight).toBe(18);
  });

  it('reads <col> + <sheetFormatPr> from openpyxl genuine fixture', async () => {
    const { readFileSync } = await import('node:fs');
    const { dirname, resolve } = await import('node:path');
    const { fileURLToPath } = await import('node:url');
    const here = dirname(fileURLToPath(import.meta.url));
    const fixturePath = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine/empty-with-styles.xlsx');
    const bytes = readFileSync(fixturePath);
    const wb = await loadWorkbook(fromBuffer(bytes));
    const ws = expectSheet(wb.sheets[0]?.sheet);
    // empty-with-styles.xlsx has <col min="1" max="1" width="10.7109375" bestFit="1" customWidth="1"/>
    const colA = getColumnDimension(ws, 1);
    expect(colA?.width).toBeCloseTo(10.7109375);
    expect(colA?.bestFit).toBe(true);
    expect(colA?.customWidth).toBe(true);
    // <sheetFormatPr defaultRowHeight="15"/>
    expect(ws.defaultRowHeight).toBe(15);
  });
});
