import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

describe('docProps + theme passthrough on load', () => {
  it('reads core / app properties and theme bytes from empty.xlsx', async () => {
    const bytes = readFileSync(resolve(FIXTURES, 'empty.xlsx'));
    const wb = await loadWorkbook(fromBuffer(bytes));
    // empty.xlsx has both core and app props.
    expect(wb.properties).toBeDefined();
    expect(wb.appProperties).toBeDefined();
    // Theme is shipped with every Excel-produced file.
    expect(wb.themeXml).toBeInstanceOf(Uint8Array);
    expect((wb.themeXml as Uint8Array).byteLength).toBeGreaterThan(0);
    // No customProperties in this fixture.
    expect(wb.customProperties).toBeUndefined();
  });
});

describe('docProps + theme passthrough through save', () => {
  it('round-trips theme bytes verbatim', async () => {
    const bytes = readFileSync(resolve(FIXTURES, 'empty.xlsx'));
    const wb = await loadWorkbook(fromBuffer(bytes));
    const themeBefore = wb.themeXml;
    expect(themeBefore).toBeDefined();

    const reSaved = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(reSaved));
    expect(wb2.themeXml).toBeDefined();
    expect(wb2.themeXml?.byteLength).toBe(themeBefore?.byteLength);
    // Theme is treated as opaque bytes — assert byte-for-byte equality.
    if (themeBefore && wb2.themeXml) {
      expect(Buffer.compare(Buffer.from(themeBefore), Buffer.from(wb2.themeXml))).toBe(0);
    }
  });

  it('round-trips core properties through save', async () => {
    const bytes = readFileSync(resolve(FIXTURES, 'empty.xlsx'));
    const wb = await loadWorkbook(fromBuffer(bytes));
    expect(wb.properties).toBeDefined();
    const reSaved = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(reSaved));
    expect(wb2.properties).toBeDefined();
    // Core property fields the openpyxl fixture sets — at minimum we have
    // creator / created / modified / lastModifiedBy in some form. We don't
    // assert exact values, just that the fields survive.
    expect(wb2.properties?.creator).toBeDefined();
  });

  it('round-trips app (extended) properties through save', async () => {
    const bytes = readFileSync(resolve(FIXTURES, 'empty.xlsx'));
    const wb = await loadWorkbook(fromBuffer(bytes));
    expect(wb.appProperties).toBeDefined();
    const reSaved = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(reSaved));
    expect(wb2.appProperties).toBeDefined();
  });

  it('omits docProps + theme when the workbook has none', async () => {
    const { createWorkbook, addWorksheet } = await import('../../src/workbook/workbook');
    const wb = createWorkbook();
    addWorksheet(wb, 'Plain');
    const reSaved = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(reSaved));
    expect(wb2.themeXml).toBeUndefined();
    expect(wb2.properties).toBeUndefined();
    expect(wb2.appProperties).toBeUndefined();
    expect(wb2.customProperties).toBeUndefined();
  });
});
