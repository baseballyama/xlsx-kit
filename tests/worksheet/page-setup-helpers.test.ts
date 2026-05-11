// Tests for the page-setup ergonomic helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import {
  addColBreak,
  addRowBreak,
  setFitToPage,
  setFooter,
  setHeader,
  setPageMargins,
  setPageOrientation,
  setPaperSize,
  setPrintScale,
} from '../../src/worksheet/page-setup';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('page-setup ergonomic helpers', () => {
  it('setPageOrientation / setPaperSize / setPrintScale + setFitToPage allocate pageSetup lazily', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(ws.pageSetup).toBeUndefined();

    setPageOrientation(ws, 'landscape');
    expect(ws.pageSetup?.orientation).toBe('landscape');
    setPaperSize(ws, 9);
    expect(ws.pageSetup?.paperSize).toBe(9);
    setPrintScale(ws, 80);
    expect(ws.pageSetup?.scale).toBe(80);
    setFitToPage(ws, { width: 1, height: 0 });
    expect(ws.pageSetup?.fitToWidth).toBe(1);
    expect(ws.pageSetup?.fitToHeight).toBe(0);
  });

  it('setPageMargins fills Excel defaults for missing axes', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setPageMargins(ws, { left: 0.5, right: 0.5 });
    expect(ws.pageMargins).toEqual({
      left: 0.5,
      right: 0.5,
      top: 1,
      bottom: 1,
      header: 0.5,
      footer: 0.5,
    });
  });

  it('setHeader / setFooter toggle differentOddEven / differentFirst as needed', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setHeader(ws, 'odd', '&CMain');
    expect(ws.headerFooter?.oddHeader).toBe('&CMain');
    expect(ws.headerFooter?.differentOddEven).toBeUndefined();
    expect(ws.headerFooter?.differentFirst).toBeUndefined();

    setHeader(ws, 'even', '&CEven');
    expect(ws.headerFooter?.evenHeader).toBe('&CEven');
    expect(ws.headerFooter?.differentOddEven).toBe(true);

    setFooter(ws, 'first', '&CFirst');
    expect(ws.headerFooter?.firstFooter).toBe('&CFirst');
    expect(ws.headerFooter?.differentFirst).toBe(true);
  });

  it('addRowBreak / addColBreak push manual breaks with sane defaults', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const r = addRowBreak(ws, 40);
    expect(r).toEqual({ id: 40, man: true, max: 16383 });
    expect(ws.rowBreaks).toEqual([r]);

    const c = addColBreak(ws, 4);
    expect(c).toEqual({ id: 4, man: true, max: 1048575 });
    expect(ws.colBreaks).toEqual([c]);
  });

  it('full save → load round-trip with helpers exercised', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setPageOrientation(ws, 'landscape');
    setPaperSize(ws, 9);
    setPrintScale(ws, 80);
    setFitToPage(ws, { width: 1, height: 0 });
    setPageMargins(ws, { left: 0.5, right: 0.5, top: 1.25, bottom: 1.25 });
    setHeader(ws, 'odd', '&CMain header');
    setFooter(ws, 'odd', '&CPage &P of &N');
    addRowBreak(ws, 50);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.pageSetup?.orientation).toBe('landscape');
    expect(ws2.pageSetup?.paperSize).toBe(9);
    expect(ws2.pageSetup?.fitToWidth).toBe(1);
    expect(ws2.pageMargins?.left).toBeCloseTo(0.5);
    expect(ws2.pageMargins?.top).toBeCloseTo(1.25);
    expect(ws2.headerFooter?.oddHeader).toBe('&CMain header');
    expect(ws2.headerFooter?.oddFooter).toBe('&CPage &P of &N');
    expect(ws2.rowBreaks[0]?.id).toBe(50);
  });
});