import { describe, expect, it } from 'vitest';
import {
  ARC_CONTENT_TYPES,
  ARC_WORKBOOK,
  CHART_NS,
  CONTYPES_NS,
  COREPROPS_NS,
  CX_NS,
  DEFAULT_PREFIXES,
  DRAWING_NS,
  EXTERNAL_LINK_NS,
  PKG_REL_NS,
  parseQName,
  qname,
  REL_NS,
  SHEET_DRAWING_NS,
  SHEET_MAIN_NS,
  XLSM_TYPE,
  XLSX_TYPE,
  XML_NS,
} from '../../../src/xml/namespaces';

describe('OOXML namespace constants', () => {
  it('match the openpyxl reference values byte-for-byte', () => {
    // Cross-checked against reference/openpyxl/openpyxl/xml/constants.py.
    expect(SHEET_MAIN_NS).toBe('http://schemas.openxmlformats.org/spreadsheetml/2006/main');
    expect(REL_NS).toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
    expect(PKG_REL_NS).toBe('http://schemas.openxmlformats.org/package/2006/relationships');
    expect(CONTYPES_NS).toBe('http://schemas.openxmlformats.org/package/2006/content-types');
    expect(COREPROPS_NS).toBe('http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
    expect(CHART_NS).toBe('http://schemas.openxmlformats.org/drawingml/2006/chart');
    expect(DRAWING_NS).toBe('http://schemas.openxmlformats.org/drawingml/2006/main');
    expect(SHEET_DRAWING_NS).toBe('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
    expect(EXTERNAL_LINK_NS).toBe(`${REL_NS}/externalLink`);
    expect(XML_NS).toBe('http://www.w3.org/XML/1998/namespace');
  });

  it('chartex (Microsoft 2014 chart) namespace is present', () => {
    expect(CX_NS).toBe('http://schemas.microsoft.com/office/drawing/2014/chartex');
  });

  it('content-type strings match the openpyxl format', () => {
    expect(XLSX_TYPE).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
    expect(XLSM_TYPE).toBe('application/vnd.ms-excel.sheet.macroEnabled.main+xml');
  });

  it('package paths match the openpyxl ARC_* values', () => {
    expect(ARC_CONTENT_TYPES).toBe('[Content_Types].xml');
    expect(ARC_WORKBOOK).toBe('xl/workbook.xml');
  });

  it('DEFAULT_PREFIXES is frozen and round-trips known prefixes', () => {
    expect(Object.isFrozen(DEFAULT_PREFIXES)).toBe(true);
    expect(DEFAULT_PREFIXES[CHART_NS]).toBe('c');
    expect(DEFAULT_PREFIXES[DRAWING_NS]).toBe('a');
    expect(DEFAULT_PREFIXES[SHEET_DRAWING_NS]).toBe('xdr');
    expect(DEFAULT_PREFIXES[REL_NS]).toBe('r');
    // SpreadsheetML main is the default (empty prefix) in normal serialisation.
    expect(DEFAULT_PREFIXES[SHEET_MAIN_NS]).toBe('');
  });
});

describe('qname / parseQName', () => {
  it('qname builds Clark notation', () => {
    expect(qname(SHEET_MAIN_NS, 'workbook')).toBe(`{${SHEET_MAIN_NS}}workbook`);
  });

  it('qname returns the bare local name when ns is empty', () => {
    expect(qname('', 'plain')).toBe('plain');
    expect(qname(undefined, 'plain')).toBe('plain');
  });

  it('parseQName splits a Clark name', () => {
    expect(parseQName(`{${SHEET_MAIN_NS}}workbook`)).toEqual({ ns: SHEET_MAIN_NS, local: 'workbook' });
  });

  it('parseQName treats unprefixed names as having an empty namespace', () => {
    expect(parseQName('plain')).toEqual({ ns: '', local: 'plain' });
  });

  it('parseQName round-trips through qname', () => {
    const original = qname(CHART_NS, 'barChart');
    const parsed = parseQName(original);
    expect(qname(parsed.ns, parsed.local)).toBe(original);
  });
});
