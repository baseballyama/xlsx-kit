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
});
