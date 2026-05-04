import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook, parseSheetEntries, resolveRelTarget } from '../../src/public/load';
import { parseXml } from '../../src/xml/parser';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

describe('resolveRelTarget', () => {
  it('treats /-prefixed targets as package-absolute', () => {
    expect(resolveRelTarget('xl/workbook.xml', '/xl/styles.xml')).toBe('xl/styles.xml');
  });

  it('joins relative targets against the source part dir', () => {
    expect(resolveRelTarget('xl/workbook.xml', 'worksheets/sheet1.xml')).toBe('xl/worksheets/sheet1.xml');
  });

  it('handles root-rels source where parent dir is empty', () => {
    expect(resolveRelTarget('', 'xl/workbook.xml')).toBe('xl/workbook.xml');
  });

  it('collapses .. segments', () => {
    expect(resolveRelTarget('xl/worksheets/sheet1.xml', '../theme/theme1.xml')).toBe('xl/theme/theme1.xml');
  });
});

describe('parseSheetEntries', () => {
  it('extracts name / sheetId / r:id from a multi-sheet workbook.xml', () => {
    const workbookXml = readFileSync(resolve(FIXTURES, 'empty.xlsx'));
    // Use the loadWorkbook flow indirectly: parse the workbook.xml inside.
    // Easier — read the raw entry via openZip elsewhere; here we just unit-test
    // parseSheetEntries against a hand-built XML string.
    const xml = `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Alpha" sheetId="1" r:id="rId1"/>
    <sheet name="Beta" sheetId="3" r:id="rId2" state="hidden"/>
    <sheet name="Gamma" sheetId="7" r:id="rId3" state="veryHidden"/>
  </sheets>
</workbook>`;
    const entries = parseSheetEntries(parseXml(xml));
    expect(entries).toEqual([
      { name: 'Alpha', sheetId: 1, rId: 'rId1', state: 'visible' },
      { name: 'Beta', sheetId: 3, rId: 'rId2', state: 'hidden' },
      { name: 'Gamma', sheetId: 7, rId: 'rId3', state: 'veryHidden' },
    ]);
    // workbookXml is loaded only to verify the fixture itself is reachable.
    expect(workbookXml.byteLength).toBeGreaterThan(0);
  });

  it('throws on a sheet missing required attrs', () => {
    expect(() =>
      parseSheetEntries(
        parseXml(
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="A" sheetId="1"/></sheets></workbook>',
        ),
      ),
    ).toThrowError(/r:id/);
  });
});

describe('loadWorkbook — empty.xlsx skeleton', () => {
  it('reads openpyxl genuine/empty.xlsx and produces a 3-sheet Workbook scaffold', async () => {
    const bytes = readFileSync(resolve(FIXTURES, 'empty.xlsx'));
    const wb = await loadWorkbook(fromBuffer(bytes));

    expect(wb.sheets.map((s) => s.sheet.title)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    expect(wb.sheets.map((s) => s.sheetId)).toEqual([1, 2, 3]);
    expect(wb.sheets.every((s) => s.state === 'visible')).toBe(true);
    // Cells haven't been read in the minimum-skeleton stage.
    for (const ref of wb.sheets) expect(ref.sheet.rows.size).toBe(0);
  });

  it('rejects an archive missing [Content_Types].xml', async () => {
    const { createZipWriter } = await import('../../src/zip/writer');
    const { toBuffer } = await import('../../src/io/node');
    const sink = toBuffer();
    const w = createZipWriter(sink);
    await w.addEntry('foo.txt', new TextEncoder().encode('bar'));
    await w.finalize();
    await expect(loadWorkbook(fromBuffer(sink.result()))).rejects.toThrow(/Content_Types/);
  });
});
