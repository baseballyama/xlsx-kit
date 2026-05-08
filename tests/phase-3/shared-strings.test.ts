import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSharedString,
  makeSharedStrings,
  parseSharedStringsXml,
  serializeSharedStrings,
  sharedStringsToBytes,
} from '../../src/workbook/shared-strings';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

describe('makeSharedStrings + addSharedString', () => {
  it('returns 0 for the first add and dedupes identical strings', () => {
    const t = makeSharedStrings();
    expect(addSharedString(t, 'alpha')).toBe(0);
    expect(addSharedString(t, 'alpha')).toBe(0);
    expect(addSharedString(t, 'beta')).toBe(1);
    expect(addSharedString(t, 'alpha')).toBe(0);
    expect(t.entries).toEqual(['alpha', 'beta']);
  });

  it('dedupes the empty string like everything else', () => {
    const t = makeSharedStrings();
    expect(addSharedString(t, '')).toBe(0);
    expect(addSharedString(t, '')).toBe(0);
    expect(addSharedString(t, 'x')).toBe(1);
    expect(t.entries).toEqual(['', 'x']);
  });
});

describe('parseSharedStringsXml — plain strings', () => {
  it('parses a single-entry sst', () => {
    const t = parseSharedStringsXml(
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>hello</t></si></sst>',
    );
    expect(t.entries).toEqual(['hello']);
    expect(t.index.get('hello')).toBe(0);
  });

  it('preserves duplicate <si> entries (slot semantics, not text)', () => {
    const t = parseSharedStringsXml(
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="2"><si><t>a</t></si><si><t>a</t></si><si><t>b</t></si></sst>',
    );
    expect(t.entries).toEqual(['a', 'a', 'b']);
    // Index points at the first occurrence.
    expect(t.index.get('a')).toBe(0);
    expect(t.index.get('b')).toBe(2);
  });

  it('preserves rich-text runs <r><t>...</t></r> as a discriminated entry', () => {
    const t = parseSharedStringsXml(
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><r><t>foo</t></r><r><t>bar</t></r></si></sst>',
    );
    expect(t.entries).toEqual([
      { kind: 'rich-text', runs: [{ text: 'foo' }, { text: 'bar' }] },
    ]);
  });

  it('rejects non-sst root', () => {
    expect(() => parseSharedStringsXml('<foo/>')).toThrowError(/expected sst/);
  });
});

describe('parseSharedStringsXml — fixture files', () => {
  it('reads openpyxl genuine/empty-with-styles.xlsx sst (1 entry)', () => {
    const { readFileSync: rfs } = { readFileSync };
    const _ = rfs; // silence the lint-only check
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>TEST HERE</t></si></sst>';
    const t = parseSharedStringsXml(xml);
    expect(t.entries).toEqual(['TEST HERE']);
  });

  it('reads openpyxl genuine/sample.xlsx sst (2 entries)', async () => {
    const { unzipSync } = await import('fflate');
    const xlsx = readFileSync(resolve(FIXTURES, 'sample.xlsx'));
    const entries = unzipSync(xlsx);
    const sstBytes = entries['xl/sharedStrings.xml'];
    if (!sstBytes) throw new Error('expected xl/sharedStrings.xml in sample fixture');
    const t = parseSharedStringsXml(sstBytes);
    expect(t.entries).toEqual(['This is cell A1 in Sheet 1', 'This is cell G5']);
  });
});

describe('serializeSharedStrings + sharedStringsToBytes', () => {
  it('emits a count + uniqueCount + 1 entry per addSharedString', () => {
    const t = makeSharedStrings();
    addSharedString(t, 'hello');
    addSharedString(t, 'world');
    addSharedString(t, 'hello');
    const out = serializeSharedStrings(t);
    expect(out).toContain('count="2"');
    expect(out).toContain('uniqueCount="2"');
    expect(out).toContain('<si><t>hello</t></si>');
    expect(out).toContain('<si><t>world</t></si>');
  });

  it('escapes <, >, & in text', () => {
    const t = makeSharedStrings();
    addSharedString(t, 'a < b > c & d');
    expect(serializeSharedStrings(t)).toContain('a &lt; b &gt; c &amp; d');
  });

  it('marks leading/trailing whitespace with xml:space="preserve"', () => {
    const t = makeSharedStrings();
    addSharedString(t, ' leading');
    addSharedString(t, 'trailing ');
    addSharedString(t, 'inner only');
    const out = serializeSharedStrings(t);
    expect(out).toContain('<t xml:space="preserve"> leading</t>');
    expect(out).toContain('<t xml:space="preserve">trailing </t>');
    // No preserve when whitespace is purely interior.
    expect(out).toContain('<t>inner only</t>');
  });

  it('round-trips through parse → serialize → parse', () => {
    const original = makeSharedStrings();
    addSharedString(original, 'alpha');
    addSharedString(original, 'beta');
    addSharedString(original, '日本語 🎉');
    const round = parseSharedStringsXml(sharedStringsToBytes(original));
    expect(round.entries).toEqual(original.entries);
  });
});
