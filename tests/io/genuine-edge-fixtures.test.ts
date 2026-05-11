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
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';

const TESTS_DATA = resolve(__dirname, '../../reference/openpyxl/openpyxl/tests/data');
const FIXTURES = TESTS_DATA;
// openpyxl mirrors a second fixture tree under `reader/tests/data` —
// some files live only in one or the other.
const READER_TESTS_DATA = resolve(__dirname, '../../reference/openpyxl/openpyxl/reader/tests/data');

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

  it('hidden_sheets.xlsx: visible / hidden / veryHidden states round-trip', async () => {
    const original = readFileSync(
      `${READER_TESTS_DATA}/hidden_sheets.xlsx`,
    );
    const wb = await loadWorkbook(fromBuffer(original));
    const states = wb.sheets.map((s) => ({ title: s.sheet.title, state: s.state }));
    expect(states).toEqual([
      { title: 'Sheet', state: 'visible' },
      { title: 'Hidden', state: 'hidden' },
      { title: 'VeryHidden', state: 'veryHidden' },
    ]);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const states2 = wb2.sheets.map((s) => ({ title: s.sheet.title, state: s.state }));
    expect(states2).toEqual(states);
  });

  it('contains_chartsheets.xlsx: chartsheet kind survives the round-trip', async () => {
    const original = readFileSync(
      `${READER_TESTS_DATA}/contains_chartsheets.xlsx`,
    );
    const wb = await loadWorkbook(fromBuffer(original));
    const kinds = wb.sheets.map((s) => s.kind);
    // The fixture has both worksheets and chartsheets; we don't assert
    // the exact mix (it may shift if openpyxl regenerates the fixture),
    // only that at least one chartsheet survives.
    expect(kinds).toContain('chartsheet');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.map((s) => s.kind)).toEqual(kinds);
  });

  it('null_file.xlsx: surfaces a clear "not a valid zip" error', async () => {
    const original = readFileSync(
      `${READER_TESTS_DATA}/null_file.xlsx`,
    );
    await expect(loadWorkbook(fromBuffer(original))).rejects.toThrowError(
      /not a valid zip|too short|EOCD/i,
    );
  });

  it('comments/tests/data/comments.xlsx: per-sheet legacyComments round-trip', async () => {
    const original = readFileSync(
      resolve(__dirname, '../../reference/openpyxl/openpyxl/comments/tests/data/comments.xlsx'),
    );
    const wb = await loadWorkbook(fromBuffer(original));
    // Sheet1 carries 6 comments, Sheet2 zero, Sheet3 one — per the
    // openpyxl reference fixture inventory.
    const counts = wb.sheets.map((s) => ({
      title: s.sheet.title,
      n: s.kind === 'worksheet' ? s.sheet.legacyComments.length : 0,
    }));
    expect(counts).toEqual([
      { title: 'Sheet1', n: 6 },
      { title: 'Sheet2', n: 0 },
      { title: 'Sheet3', n: 1 },
    ]);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const counts2 = wb2.sheets.map((s) => ({
      title: s.sheet.title,
      n: s.kind === 'worksheet' ? s.sheet.legacyComments.length : 0,
    }));
    expect(counts2).toEqual(counts);
  });

  it('packaging/tests/data/hyperlink.xlsx: external URL round-trip', async () => {
    const original = readFileSync(
      resolve(__dirname, '../../reference/openpyxl/openpyxl/packaging/tests/data/hyperlink.xlsx'),
    );
    const wb = await loadWorkbook(fromBuffer(original));
    const s0 = wb.sheets[0];
    if (s0?.kind !== 'worksheet') throw new Error('expected worksheet');
    expect(s0.sheet.hyperlinks).toHaveLength(1);
    const link = s0.sheet.hyperlinks[0];
    expect(link?.ref).toBe('A1');
    expect(link?.target).toBe('http://www.readthedocs.org');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const s0b = wb2.sheets[0];
    if (s0b?.kind !== 'worksheet') throw new Error('expected worksheet');
    expect(s0b.sheet.hyperlinks[0]?.target).toBe('http://www.readthedocs.org');
  });

  it('worksheet/tests/data/test_datetime.xlsx: numeric date serials round-trip', async () => {
    const original = readFileSync(
      resolve(__dirname, '../../reference/openpyxl/openpyxl/worksheet/tests/data/test_datetime.xlsx'),
    );
    const wb = await loadWorkbook(fromBuffer(original));
    expect(wb.date1904).toBe(false);
    const s0 = wb.sheets[0];
    if (s0?.kind !== 'worksheet') throw new Error('expected worksheet');
    // First cell is the openpyxl test datetime serial (≈ 2.0987 days).
    expect(s0.sheet.rows.get(1)?.get(1)?.value).toBeGreaterThan(0);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const s0b = wb2.sheets[0];
    if (s0b?.kind !== 'worksheet') throw new Error('expected worksheet');
    expect(s0b.sheet.rows.get(1)?.get(1)?.value).toEqual(s0.sheet.rows.get(1)?.get(1)?.value);
  });

  it('contains_chartsheets.xlsx: xl/calcChain.xml passthrough preserves the byte-identical entry', async () => {
    const original = readFileSync(`${READER_TESTS_DATA}/contains_chartsheets.xlsx`);
    const wb = await loadWorkbook(fromBuffer(original));
    expect(wb.passthrough?.has('xl/calcChain.xml')).toBe(true);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.passthrough?.get('xl/calcChain.xml')).toEqual(wb.passthrough?.get('xl/calcChain.xml'));
  });

  it('reader/legacy_drawing.xlsm: control-VML + ctrlProps survive', async () => {
    const original = readFileSync(`${READER_TESTS_DATA}/legacy_drawing.xlsm`);
    const wb = await loadWorkbook(fromBuffer(original));
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    // Both round-trips end up with the same ctrlProps + VML set.
    const before = [...(wb.passthrough?.keys() ?? [])].sort();
    const after = [...(wb2.passthrough?.keys() ?? [])].sort();
    expect(after).toEqual(before);
    expect(wb.sheets.length).toBe(wb2.sheets.length);
  });
});
