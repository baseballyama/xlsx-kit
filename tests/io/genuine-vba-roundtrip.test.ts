// Phase 7 §3 acceptance: real openpyxl xlsm fixture round-trip. — VBA を含む xlsm
// の round-trip で vbaProject.bin が byte-identical, ActiveX / ctrlProps /
// customUI が 消えない、keepVba: false で xlsm を読むと VBA が消えて xlsx になる。

import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';

const FIXTURE = resolve(
  __dirname,
  '../../reference/openpyxl/openpyxl/tests/data/reader/vba+comments.xlsm',
);

describe('Phase 7 — genuine xlsm round-trip (openpyxl vba+comments.xlsm)', () => {
  it('loads → saves → reloads with vbaProject.bin byte-identical', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));

    expect(wb.vbaProject).toBeDefined();
    expect(wb.vbaProject?.length).toBe(14848);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.vbaProject).toEqual(wb.vbaProject);
  });

  it('captures all 10 ctrlProps as passthrough and round-trips them', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));

    const ctrlPropPaths = [...(wb.passthrough?.keys() ?? [])]
      .filter((p) => p.startsWith('xl/ctrlProps/'))
      .sort();
    expect(ctrlPropPaths).toEqual([
      'xl/ctrlProps/ctrlProp1.xml',
      'xl/ctrlProps/ctrlProp10.xml',
      'xl/ctrlProps/ctrlProp2.xml',
      'xl/ctrlProps/ctrlProp3.xml',
      'xl/ctrlProps/ctrlProp4.xml',
      'xl/ctrlProps/ctrlProp5.xml',
      'xl/ctrlProps/ctrlProp6.xml',
      'xl/ctrlProps/ctrlProp7.xml',
      'xl/ctrlProps/ctrlProp8.xml',
      'xl/ctrlProps/ctrlProp9.xml',
    ]);

    // Round-trip survives — all 10 ctrlProps still in the passthrough Map.
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ctrlPropPaths2 = [...(wb2.passthrough?.keys() ?? [])]
      .filter((p) => p.startsWith('xl/ctrlProps/'))
      .sort();
    expect(ctrlPropPaths2).toEqual(ctrlPropPaths);

    // And bytes match.
    for (const p of ctrlPropPaths) {
      expect(wb2.passthrough?.get(p)).toEqual(wb.passthrough?.get(p));
    }
  });

  it('captures printerSettings binary as passthrough', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));

    const printer = wb.passthrough?.get('xl/printerSettings/printerSettings1.bin');
    expect(printer).toBeDefined();
    expect(printer?.length).toBe(1040);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.passthrough?.get('xl/printerSettings/printerSettings1.bin')).toEqual(printer);
  });

  it('promotes the workbook content type to xlsm on round-trip save', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));
    const bytes = await workbookToBytes(wb);

    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    const ct = new TextDecoder().decode(entries['[Content_Types].xml']);
    expect(ct).toContain('macroEnabled.main+xml');
  });

  it('preserves the worksheet structure across the round-trip', async () => {
    const original = readFileSync(FIXTURE);
    const wb = await loadWorkbook(fromBuffer(original));
    expect(wb.sheets.length).toBeGreaterThanOrEqual(1);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.length).toBe(wb.sheets.length);
    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(wb.sheets.map((s) => s.sheet.title));
  });
});
