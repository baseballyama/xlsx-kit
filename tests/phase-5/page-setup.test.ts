// Tests for the typed page-setup model. Per docs/plan/13-full-excel-coverage.md
// §B6 (printOptions / pageMargins / pageSetup / headerFooter).

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  makeHeaderFooter,
  makePageMargins,
  makePageSetup,
  makePrintOptions,
} from '../../src/worksheet/page-setup';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('printOptions round-trip', () => {
  it('preserves the 5 boolean flags', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'P');
    setCell(ws, 1, 1, 1);
    ws.printOptions = makePrintOptions({
      horizontalCentered: true,
      verticalCentered: false,
      headings: true,
      gridLines: true,
      gridLinesSet: false,
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.printOptions?.horizontalCentered).toBe(true);
    expect(ws2.printOptions?.verticalCentered).toBe(false);
    expect(ws2.printOptions?.headings).toBe(true);
    expect(ws2.printOptions?.gridLines).toBe(true);
    expect(ws2.printOptions?.gridLinesSet).toBe(false);
  });
});

describe('pageMargins round-trip', () => {
  it('preserves all six margins', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'M');
    setCell(ws, 1, 1, 1);
    ws.pageMargins = makePageMargins({ left: 0.5, right: 0.5, top: 1.25, bottom: 1.25, header: 0.3, footer: 0.3 });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.pageMargins?.left).toBeCloseTo(0.5);
    expect(ws2.pageMargins?.right).toBeCloseTo(0.5);
    expect(ws2.pageMargins?.top).toBeCloseTo(1.25);
    expect(ws2.pageMargins?.bottom).toBeCloseTo(1.25);
    expect(ws2.pageMargins?.header).toBeCloseTo(0.3);
    expect(ws2.pageMargins?.footer).toBeCloseTo(0.3);
  });
});

describe('pageSetup round-trip', () => {
  it('preserves orientation / paperSize / fitToWidth / DPI', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 1);
    ws.pageSetup = makePageSetup({
      paperSize: 9,
      orientation: 'landscape',
      scale: 80,
      fitToWidth: 1,
      fitToHeight: 0,
      pageOrder: 'overThenDown',
      horizontalDpi: 300,
      verticalDpi: 300,
      blackAndWhite: true,
      cellComments: 'asDisplayed',
      errors: 'NA',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const ps = ws2.pageSetup;
    expect(ps?.paperSize).toBe(9);
    expect(ps?.orientation).toBe('landscape');
    expect(ps?.scale).toBe(80);
    expect(ps?.fitToWidth).toBe(1);
    expect(ps?.fitToHeight).toBe(0);
    expect(ps?.pageOrder).toBe('overThenDown');
    expect(ps?.horizontalDpi).toBe(300);
    expect(ps?.verticalDpi).toBe(300);
    expect(ps?.blackAndWhite).toBe(true);
    expect(ps?.cellComments).toBe('asDisplayed');
    expect(ps?.errors).toBe('NA');
  });
});

describe('headerFooter round-trip', () => {
  it('preserves odd/first headers and footers + flags', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'H');
    setCell(ws, 1, 1, 1);
    ws.headerFooter = makeHeaderFooter({
      differentFirst: true,
      differentOddEven: true,
      scaleWithDoc: false,
      alignWithMargins: false,
      oddHeader: '&LLeft&CCenter — &P / &N&R&D',
      oddFooter: '&L&F&CPage &P&RConfidential',
      firstHeader: '&CFirst Page Only',
      evenHeader: '&CEven Pages',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const hf = ws2.headerFooter;
    expect(hf?.differentFirst).toBe(true);
    expect(hf?.differentOddEven).toBe(true);
    expect(hf?.scaleWithDoc).toBe(false);
    expect(hf?.alignWithMargins).toBe(false);
    expect(hf?.oddHeader).toBe('&LLeft&CCenter — &P / &N&R&D');
    expect(hf?.oddFooter).toBe('&L&F&CPage &P&RConfidential');
    expect(hf?.firstHeader).toBe('&CFirst Page Only');
    expect(hf?.evenHeader).toBe('&CEven Pages');
    expect(hf?.firstFooter).toBeUndefined();
    expect(hf?.evenFooter).toBeUndefined();
  });
});

describe('default emission behavior', () => {
  it('emits no page-setup elements when none are set', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'N');
    setCell(ws, 1, 1, 'a');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.printOptions).toBeUndefined();
    expect(ws2.pageMargins).toBeUndefined();
    expect(ws2.pageSetup).toBeUndefined();
    expect(ws2.headerFooter).toBeUndefined();
  });
});