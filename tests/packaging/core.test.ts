import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { corePropsFromBytes, corePropsToBytes, makeCoreProperties } from '../../src/packaging/core';
import { openZip } from '../../src/zip/reader';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

describe('coreProperties — basic operations', () => {
  it('makeCoreProperties yields an empty object', () => {
    expect(makeCoreProperties()).toEqual({});
  });

  it('round-trips a small hand-built CoreProperties', () => {
    const p = {
      title: 'Quarterly Report',
      creator: 'Yuichiro',
      lastModifiedBy: 'Yuichiro',
      created: '2026-01-15T09:00:00Z',
      modified: '2026-05-04T13:30:00Z',
      keywords: 'budget,q1',
    };
    const bytes = corePropsToBytes(p);
    const back = corePropsFromBytes(bytes);
    expect(back).toEqual(p);
  });

  it('emits xsi:type="dcterms:W3CDTF" on dcterms timestamps', () => {
    const p = { created: '2026-01-15T09:00:00Z' };
    const xml = new TextDecoder().decode(corePropsToBytes(p));
    expect(xml).toContain('xsi:type="dcterms:W3CDTF"');
    expect(xml).toContain('<dcterms:created');
    expect(xml).toContain('2026-01-15T09:00:00Z');
  });

  it('omits absent fields from the XML entirely', () => {
    const p = { creator: 'A' };
    const xml = new TextDecoder().decode(corePropsToBytes(p));
    expect(xml).toContain('<dc:creator>A</dc:creator>');
    expect(xml).not.toContain('<dc:title');
    expect(xml).not.toContain('<dcterms:modified');
  });
});

describe('coreProperties — openpyxl fixture round-trip', () => {
  it('parses docProps/core.xml from genuine/empty.xlsx', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const p = corePropsFromBytes(zip.read('docProps/core.xml'));
    // Cross-checked against the live XML inspected this turn:
    // <dc:creator>CED501</dc:creator>
    // <cp:lastModifiedBy>CED501</cp:lastModifiedBy>
    // <dcterms:created>2010-07-28T08:40:37Z</dcterms:created>
    // <dcterms:modified>2010-07-28T08:40:56Z</dcterms:modified>
    expect(p.creator).toBe('CED501');
    expect(p.lastModifiedBy).toBe('CED501');
    expect(p.created).toBe('2010-07-28T08:40:37Z');
    expect(p.modified).toBe('2010-07-28T08:40:56Z');
  });

  it('parsing → re-serialising → re-parsing yields equal objects', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const original = zip.read('docProps/core.xml');
    const a = corePropsFromBytes(original);
    const re = corePropsToBytes(a);
    const b = corePropsFromBytes(re);
    expect(b).toEqual(a);
  });
});
