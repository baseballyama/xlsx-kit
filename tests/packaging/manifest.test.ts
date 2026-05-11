import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import {
  addDefault,
  addOverride,
  findOverride,
  findOverrideByContentType,
  makeManifest,
  manifestFromBytes,
  manifestToBytes,
} from '../../src/packaging/manifest';
import { openZip } from '../../src/zip/reader';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

describe('manifest — basic operations', () => {
  it('makeManifest yields an empty manifest', () => {
    expect(makeManifest()).toEqual({ defaults: [], overrides: [] });
  });

  it('addDefault appends and is idempotent on (ext, contentType)', () => {
    const m = makeManifest();
    addDefault(m, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');
    addDefault(m, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');
    expect(m.defaults).toHaveLength(1);
  });

  it('addDefault for an existing ext but new contentType updates in place', () => {
    const m = makeManifest();
    addDefault(m, 'xml', 'application/xml');
    addDefault(m, 'xml', 'application/something-else');
    expect(m.defaults).toEqual([{ ext: 'xml', contentType: 'application/something-else' }]);
  });

  it('addOverride / findOverride round-trip', () => {
    const m = makeManifest();
    addOverride(m, '/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
    expect(findOverride(m, '/xl/workbook.xml')?.contentType).toContain('spreadsheetml.sheet.main');
    expect(findOverride(m, '/xl/no.xml')).toBeUndefined();
  });

  it('findOverrideByContentType returns the first match', () => {
    const m = makeManifest();
    addOverride(m, '/a.xml', 'type-a');
    addOverride(m, '/b.xml', 'type-b');
    expect(findOverrideByContentType(m, 'type-b')?.partName).toBe('/b.xml');
  });
});

describe('manifest — XML round-trip', () => {
  it('manifestToBytes / manifestFromBytes round-trip a hand-built manifest', () => {
    const m = makeManifest();
    addDefault(m, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');
    addDefault(m, 'xml', 'application/xml');
    addOverride(m, '/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
    addOverride(m, '/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml');
    const bytes = manifestToBytes(m);
    const back = manifestFromBytes(bytes);
    expect(back).toEqual(m);
  });

  it('parses openpyxl genuine/empty.xlsx [Content_Types].xml without loss', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const original = zip.read('[Content_Types].xml');
    const m = manifestFromBytes(original);

    // The fixture defines `rels` and `xml` defaults plus several overrides
    // for theme/styles/workbook/sheets/docProps; cross-checked against
    // unzip -p output captured during turn 7.
    expect(m.defaults).toContainEqual({
      ext: 'rels',
      contentType: 'application/vnd.openxmlformats-package.relationships+xml',
    });
    expect(m.defaults).toContainEqual({ ext: 'xml', contentType: 'application/xml' });
    expect(findOverride(m, '/xl/workbook.xml')?.contentType).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    );
    expect(m.overrides.length).toBeGreaterThanOrEqual(7);
  });

  it('parsing → re-serialising → re-parsing yields an equivalent manifest', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const original = zip.read('[Content_Types].xml');
    const m1 = manifestFromBytes(original);
    const re = manifestToBytes(m1);
    const m2 = manifestFromBytes(re);
    expect(m2).toEqual(m1);
  });
});
