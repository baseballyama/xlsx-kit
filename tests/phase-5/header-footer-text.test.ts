// Tests for buildHeaderFooterText / setHeaderText / setFooterText +
// the HEADER_FOOTER_CODES catalogue.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import {
  buildHeaderFooterText,
  HEADER_FOOTER_CODES,
  setFooterText,
  setHeaderText,
} from '../../src/worksheet/page-setup';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import type { Worksheet } from '../../src/worksheet/worksheet';

describe('buildHeaderFooterText', () => {
  it('all three sections', () => {
    expect(
      buildHeaderFooterText({ left: 'L', center: 'C', right: 'R' }),
    ).toBe('&LL&CC&RR');
  });

  it('center-only does not prefix &L', () => {
    expect(buildHeaderFooterText({ center: 'Page' })).toBe('&CPage');
  });

  it('omits sections with undefined fragments but keeps empty strings', () => {
    expect(buildHeaderFooterText({ left: '', right: 'R' })).toBe('&L&RR');
  });

  it('empty input → empty string', () => {
    expect(buildHeaderFooterText({})).toBe('');
  });

  it('common code interpolation reads naturally', () => {
    const text = buildHeaderFooterText({
      left: HEADER_FOOTER_CODES.sheetName,
      center: `Page ${HEADER_FOOTER_CODES.pageNumber} of ${HEADER_FOOTER_CODES.pageCount}`,
      right: HEADER_FOOTER_CODES.date,
    });
    expect(text).toBe('&L&A&CPage &P of &N&R&D');
  });
});

describe('setHeaderText', () => {
  it("default section 'odd' writes oddHeader", () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setHeaderText(ws, { left: 'foo', right: 'bar' });
    expect(ws.headerFooter?.oddHeader).toBe('&Lfoo&Rbar');
    expect(ws.headerFooter?.evenHeader).toBeUndefined();
  });

  it("section 'first' flips differentFirst", () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setHeaderText(ws, { center: 'first' }, 'first');
    expect(ws.headerFooter?.firstHeader).toBe('&Cfirst');
    expect(ws.headerFooter?.differentFirst).toBe(true);
  });

  it("section 'even' flips differentOddEven", () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setHeaderText(ws, { center: 'even' }, 'even');
    expect(ws.headerFooter?.evenHeader).toBe('&Ceven');
    expect(ws.headerFooter?.differentOddEven).toBe(true);
  });
});

describe('setFooterText', () => {
  it('writes oddFooter via the same builder', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setFooterText(ws, {
      left: HEADER_FOOTER_CODES.sheetName,
      right: `${HEADER_FOOTER_CODES.pageNumber}/${HEADER_FOOTER_CODES.pageCount}`,
    });
    expect(ws.headerFooter?.oddFooter).toBe('&L&A&R&P/&N');
  });
});

describe('header/footer round-trip', () => {
  it('header text survives saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'HF');
    setHeaderText(ws, { left: 'Report', center: HEADER_FOOTER_CODES.pageNumber, right: HEADER_FOOTER_CODES.date });
    setFooterText(ws, { center: `${HEADER_FOOTER_CODES.fileName} - ${HEADER_FOOTER_CODES.sheetName}` });
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb2.sheets[0]?.sheet;
    if (!sheet || !('rows' in sheet)) throw new Error('expected worksheet');
    const ws2 = sheet as Worksheet;
    expect(ws2.headerFooter?.oddHeader).toBe('&LReport&C&P&R&D');
    expect(ws2.headerFooter?.oddFooter).toBe('&C&F - &A');
  });
});
