// Phase 3 acceptance — extra edge-case fixtures from openpyxl's
// reference corpus. Complements `genuine-roundtrip.test.ts` with
// shapes the basic empty.xlsx / sample.xlsx / empty-with-styles.xlsx
// don't exercise:
//
// - mac_date.xlsx              — `<workbookPr date1904="true">` epoch
// - libreoffice_nrt.xlsx       — LibreOffice-emitted xlsx (different
//                                 element ordering, attribute defaults)
// - nonstandard_workbook_name  — workbook part at xl/workbook10.xml
//                                 (root rels resolves to a non-default
//                                 path)
// - bigfoot.xlsx               — multi-sheet (~30 sheets) for scale

import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';

const FIXTURES = resolve(__dirname, '../../reference/openpyxl/openpyxl/tests/data');

describe('phase-3 — additional genuine fixture round-trips', () => {
  it('mac_date.xlsx: parses workbookPr@date1904 and round-trips the flag', async () => {
    const original = readFileSync(`${FIXTURES}/genuine/mac_date.xlsx`);
    const wb = await loadWorkbook(fromBuffer(original));
    expect(wb.date1904).toBe(true);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.date1904).toBe(true);
  });

  it('libreoffice_nrt.xlsx: round-trip preserves sheet content', async () => {
    const original = readFileSync(`${FIXTURES}/genuine/libreoffice_nrt.xlsx`);
    const wb = await loadWorkbook(fromBuffer(original));
    expect(wb.sheets.length).toBeGreaterThanOrEqual(1);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.length).toBe(wb.sheets.length);
    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(wb.sheets.map((s) => s.sheet.title));
  });

  it('nonstandard_workbook_name.xlsx: handles xl/workbook10.xml via root rels', async () => {
    const original = readFileSync(`${FIXTURES}/reader/nonstandard_workbook_name.xlsx`);
    const wb = await loadWorkbook(fromBuffer(original));
    expect(wb.sheets.length).toBeGreaterThanOrEqual(1);

    // Round-trip: our writer always emits xl/workbook.xml, regardless
    // of the original path. That's a deliberate normalisation; verify
    // the reload still works.
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(wb.sheets.map((s) => s.sheet.title));
  });

  it('bigfoot.xlsx: round-trip preserves all 30+ worksheets', async () => {
    const original = readFileSync(`${FIXTURES}/reader/bigfoot.xlsx`);
    const wb = await loadWorkbook(fromBuffer(original));
    // bigfoot.xlsx has 30+ sheets per the reference fixture inventory.
    expect(wb.sheets.length).toBeGreaterThanOrEqual(30);
    const titles = wb.sheets.map((s) => s.sheet.title);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(titles);
  });
});
