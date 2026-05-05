// Phase 7 §3.3 acceptance — real openpyxl pivot fixture round-trip.
// Per docs/plan/09-pivot-vba.md §2: pivot は schema 化せず passthrough。
// pivotCache / pivotTables 配下のすべてのバイトと _rels が round-trip する
// ことだけを保証し、編集 API は提供しない。

import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';

const FIXTURE = resolve(__dirname, '../../reference/openpyxl/openpyxl/reader/tests/data/pivot.xlsx');

describe('Phase 7 — genuine pivot round-trip (openpyxl pivot.xlsx)', () => {
  it('captures pivotCache + pivotTables parts and rels into passthrough', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));

    const pivotPaths = [...(wb.passthrough?.keys() ?? [])]
      .filter((p) => p.startsWith('xl/pivotCache/') || p.startsWith('xl/pivotTables/'))
      .sort();
    expect(pivotPaths).toEqual([
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
      'xl/pivotCache/pivotCacheDefinition1.xml',
      'xl/pivotCache/pivotCacheRecords1.xml',
      'xl/pivotTables/_rels/pivotTable1.xml.rels',
      'xl/pivotTables/pivotTable1.xml',
    ]);
  });

  it('round-trips pivotCacheDefinition / pivotCacheRecords / pivotTable byte-identical', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));

    const targets = [
      'xl/pivotCache/pivotCacheDefinition1.xml',
      'xl/pivotCache/pivotCacheRecords1.xml',
      'xl/pivotTables/pivotTable1.xml',
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
      'xl/pivotTables/_rels/pivotTable1.xml.rels',
    ] as const;

    const before = new Map(targets.map((p) => [p, wb.passthrough?.get(p)]));
    for (const p of targets) expect(before.get(p)).toBeDefined();

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));

    for (const p of targets) {
      expect(wb2.passthrough?.get(p), `byte mismatch on ${p}`).toEqual(before.get(p));
    }
  });

  it('preserves pivot Override content types in [Content_Types].xml', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));
    const bytes = await workbookToBytes(wb);

    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    const ct = new TextDecoder().decode(entries['[Content_Types].xml']);
    expect(ct).toContain('pivotCacheDefinition+xml');
    expect(ct).toContain('pivotCacheRecords+xml');
    expect(ct).toContain('pivotTable+xml');
  });

  it('preserves both worksheets across the round-trip', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));
    expect(wb.sheets.map((s) => s.sheet.title)).toEqual(['ptsheet', 'raw']);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(['ptsheet', 'raw']);
  });

  it('preserves <pivotCaches> in workbook.xml + matching workbook-rels entry', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));
    const bytes = await workbookToBytes(wb);

    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    const wbXml = new TextDecoder().decode(entries['xl/workbook.xml']);
    const wbRels = new TextDecoder().decode(entries['xl/_rels/workbook.xml.rels']);

    // The <pivotCaches><pivotCache cacheId="68" r:id="..."/> survives.
    const pivotCacheMatch = wbXml.match(/<pivotCache[^/]*cacheId="68"[^/]*r:id="(rId\d+)"\s*\/>/);
    expect(pivotCacheMatch, `workbook.xml should keep <pivotCache>: ${wbXml}`).not.toBeNull();
    const pivotRId = pivotCacheMatch?.[1];
    expect(pivotRId).toBeDefined();

    // …and the matching workbook-rels entry uses the same Id and points
    // at the captured pivotCacheDefinition1.xml part.
    const relPattern = new RegExp(
      `<Relationship[^/]*Id="${pivotRId}"[^/]*Type="[^"]*pivotCacheDefinition"[^/]*Target="pivotCache/pivotCacheDefinition1\\.xml"\\s*/>`,
    );
    expect(wbRels).toMatch(relPattern);
  });

  it('preserves the sheet1 → pivotTable rel chain (worksheet relsExtras)', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));

    // sheet1 ("ptsheet") originally points at xl/pivotTables/pivotTable1.xml
    // through xl/worksheets/_rels/sheet1.xml.rels. Capture surfaces this
    // on Worksheet.relsExtras.
    const sheet1 = wb.sheets[0];
    expect(sheet1?.kind).toBe('worksheet');
    if (sheet1?.kind !== 'worksheet') return;
    const pivotRel = sheet1.sheet.relsExtras?.find((r) => r.type.endsWith('/pivotTable'));
    expect(pivotRel).toBeDefined();
    expect(pivotRel?.target).toBe('../pivotTables/pivotTable1.xml');

    // Round-trip: the rels file emerges with the pivotTable rel preserved.
    const bytes = await workbookToBytes(wb);
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    const sheet1Rels = new TextDecoder().decode(entries['xl/worksheets/_rels/sheet1.xml.rels']);
    expect(sheet1Rels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"');
    expect(sheet1Rels).toContain('Target="../pivotTables/pivotTable1.xml"');
  });

  it('round-trips worksheet pageMargins + bodyExtras extLst through reload', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));

    const sheet1Round = wb2.sheets[0];
    if (sheet1Round?.kind !== 'worksheet') throw new Error('expected worksheet');
    // pageMargins now lands on the typed field (B6); extLst stays in bodyExtras.
    expect(sheet1Round.sheet.pageMargins).toBeDefined();
    const namesAfter = (sheet1Round.sheet.bodyExtras?.afterSheetData ?? []).map((n) =>
      n.name.replace(/^\{[^}]+\}/, ''),
    );
    expect(namesAfter).toContain('extLst');
  });

  it('preserves worksheet pageMargins + bodyExtras extLst on sheet1', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));

    // sheet1 has <pageMargins/> (now typed) and <extLst><ext><mx:PLV/></ext></extLst>
    // (still captured into Worksheet.bodyExtras.afterSheetData).
    const sheet1 = wb.sheets[0];
    expect(sheet1?.kind).toBe('worksheet');
    if (sheet1?.kind !== 'worksheet') return;
    expect(sheet1.sheet.pageMargins).toBeDefined();
    const after = sheet1.sheet.bodyExtras?.afterSheetData ?? [];
    const localNames = after.map((n) => n.name.replace(/^\{[^}]+\}/, ''));
    expect(localNames).toContain('extLst');

    const bytes = await workbookToBytes(wb);
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    const sheet1Xml = new TextDecoder().decode(entries['xl/worksheets/sheet1.xml']);
    expect(sheet1Xml).toContain('<pageMargins');
    expect(sheet1Xml).toContain('<extLst');
    // mx:PLV survives via captured XmlNode tree (prefix may be reallocated).
    expect(sheet1Xml).toMatch(/PLV[^/]*Mode="0"/);
  });

  it('preserves the workbook-extras (fileVersion / workbookPr / bookViews / calcPr / extLst)', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));
    const bytes = await workbookToBytes(wb);

    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    const wbXml = new TextDecoder().decode(entries['xl/workbook.xml']);

    // The before-sheets and after-sheets extras both round-trip, so Excel
    // sees the same workbook-level metadata it emitted.
    expect(wbXml).toContain('<fileVersion');
    expect(wbXml).toContain('<workbookPr');
    expect(wbXml).toContain('<bookViews');
    expect(wbXml).toContain('<workbookView');
    expect(wbXml).toContain('<calcPr');
    expect(wbXml).toContain('<extLst');
    // pivotCaches lives in afterSheets — it's the actual point of this fixture.
    expect(wbXml).toContain('<pivotCaches');
    expect(wbXml).toContain('<pivotCache');
  });
});
