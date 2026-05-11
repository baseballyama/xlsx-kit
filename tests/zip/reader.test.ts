import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { OpenXmlIoError } from '../../src/utils/exceptions';
import { openZip } from '../../src/zip/reader';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');
const EMPTY_XLSX = resolve(FIXTURES, 'empty.xlsx');

// Sizes verified against `unzip -l` on the openpyxl fixture (turn 7 plan).
const EMPTY_ENTRIES: ReadonlyArray<readonly [path: string, size: number]> = [
  ['[Content_Types].xml', 1304],
  ['_rels/.rels', 588],
  ['docProps/app.xml', 850],
  ['docProps/core.xml', 605],
  ['xl/_rels/workbook.xml.rels', 839],
  ['xl/styles.xml', 868],
  ['xl/theme/theme1.xml', 6995],
  ['xl/workbook.xml', 624],
  ['xl/worksheets/sheet1.xml', 455],
  ['xl/worksheets/sheet2.xml', 439],
  ['xl/worksheets/sheet3.xml', 439],
];

describe('openZip(empty.xlsx)', () => {
  it('lists every expected entry, sorted', async () => {
    const zip = await openZip(fromBuffer(readFileSync(EMPTY_XLSX)));
    expect(zip.list()).toEqual(EMPTY_ENTRIES.map(([p]) => p));
  });

  it('inflates each entry to the expected uncompressed size', async () => {
    const zip = await openZip(fromBuffer(readFileSync(EMPTY_XLSX)));
    for (const [path, size] of EMPTY_ENTRIES) {
      const bytes = zip.read(path);
      expect(bytes).toBeInstanceOf(Uint8Array);
      expect(bytes.byteLength).toBe(size);
    }
  });

  it('readAsync yields the same bytes as read', async () => {
    const zip = await openZip(fromBuffer(readFileSync(EMPTY_XLSX)));
    const sync = zip.read('xl/workbook.xml');
    const async = await zip.readAsync('xl/workbook.xml');
    expect(async).toEqual(sync);
  });

  it('has() returns true only for present entries', async () => {
    const zip = await openZip(fromBuffer(readFileSync(EMPTY_XLSX)));
    expect(zip.has('xl/workbook.xml')).toBe(true);
    expect(zip.has('xl/no-such-file.xml')).toBe(false);
  });

  it('OOXML headers are present in the inflated workbook part', async () => {
    const zip = await openZip(fromBuffer(readFileSync(EMPTY_XLSX)));
    const workbookXml = new TextDecoder().decode(zip.read('xl/workbook.xml'));
    expect(workbookXml.startsWith('<?xml')).toBe(true);
    expect(workbookXml).toContain('<workbook');
  });

  it('throws OpenXmlIoError for an unknown path', async () => {
    const zip = await openZip(fromBuffer(readFileSync(EMPTY_XLSX)));
    expect(() => zip.read('xl/nope.xml')).toThrowError(OpenXmlIoError);
  });

  it('throws OpenXmlIoError for non-zip input', async () => {
    const garbage = new TextEncoder().encode('not a zip');
    await expect(openZip(fromBuffer(garbage))).rejects.toBeInstanceOf(OpenXmlIoError);
  });

  it('close() releases the entry table and subsequent reads throw', async () => {
    const zip = await openZip(fromBuffer(readFileSync(EMPTY_XLSX)));
    zip.close();
    expect(() => zip.read('xl/workbook.xml')).toThrowError(OpenXmlIoError);
    expect(() => zip.list()).toThrowError(OpenXmlIoError);
    expect(zip.has('xl/workbook.xml')).toBe(false);
  });
});
