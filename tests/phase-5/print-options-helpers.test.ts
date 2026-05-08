// Tests for print-option ergonomic helpers (setPrintGridLines /
// setPrintHeadings / setPrintCentered).

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import {
  setPrintCentered,
  setPrintGridLines,
  setPrintHeadings,
} from '../../src/xlsx/worksheet/page-setup';
import { loadWorkbook } from '../../src/xlsx/io/load';
import { workbookToBytes } from '../../src/xlsx/io/save';
import { addWorksheet, createWorkbook } from '../../src/xlsx/workbook/workbook';
import type { Worksheet } from '../../src/xlsx/worksheet/worksheet';

describe('setPrintGridLines', () => {
  it('lazily creates printOptions and pairs gridLines + gridLinesSet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setPrintGridLines(ws, true);
    expect(ws.printOptions?.gridLines).toBe(true);
    expect(ws.printOptions?.gridLinesSet).toBe(true);
  });

  it('flipping to false keeps both flags in sync', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setPrintGridLines(ws, true);
    setPrintGridLines(ws, false);
    expect(ws.printOptions?.gridLines).toBe(false);
    expect(ws.printOptions?.gridLinesSet).toBe(false);
  });
});

describe('setPrintHeadings', () => {
  it('toggles headings independently of gridLines', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setPrintHeadings(ws, true);
    expect(ws.printOptions?.headings).toBe(true);
    expect(ws.printOptions?.gridLines).toBeUndefined();
  });
});

describe('setPrintCentered', () => {
  it('horizontal only leaves vertical untouched', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setPrintCentered(ws, { horizontal: true });
    expect(ws.printOptions?.horizontalCentered).toBe(true);
    expect(ws.printOptions?.verticalCentered).toBeUndefined();
  });

  it('both axes set in one call', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setPrintCentered(ws, { horizontal: true, vertical: true });
    expect(ws.printOptions?.horizontalCentered).toBe(true);
    expect(ws.printOptions?.verticalCentered).toBe(true);
  });

  it('updates partially without clobbering the other axis', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setPrintCentered(ws, { horizontal: true, vertical: true });
    setPrintCentered(ws, { vertical: false });
    expect(ws.printOptions?.horizontalCentered).toBe(true);
    expect(ws.printOptions?.verticalCentered).toBe(false);
  });
});

describe('print options round-trip', () => {
  it('all 4 fields survive saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'PO');
    setPrintGridLines(ws, true);
    setPrintHeadings(ws, true);
    setPrintCentered(ws, { horizontal: true, vertical: true });
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb2.sheets[0]?.sheet;
    if (!sheet || !('rows' in sheet)) throw new Error('expected worksheet');
    const ws2 = sheet as Worksheet;
    expect(ws2.printOptions?.gridLines).toBe(true);
    expect(ws2.printOptions?.headings).toBe(true);
    expect(ws2.printOptions?.horizontalCentered).toBe(true);
    expect(ws2.printOptions?.verticalCentered).toBe(true);
  });
});
