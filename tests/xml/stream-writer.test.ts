import { describe, expect, it } from 'vitest';
import { OpenXmlIoError } from '../../src/utils/exceptions';
import { CHART_NS, REL_NS, SHEET_MAIN_NS, XML_NS } from '../../src/xml/namespaces';
import { parseXml } from '../../src/xml/parser';
import { createXmlStreamWriter } from '../../src/xml/stream-writer';
import { el } from '../../src/xml/tree';

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

describe('createXmlStreamWriter — minimal output', () => {
  it('emits the XML declaration by default and a self-closing root', () => {
    const w = createXmlStreamWriter();
    w.start('r');
    w.end();
    expect(decode(w.result())).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<r/>');
  });

  it('switches to a closing tag when text or children are present', () => {
    const w = createXmlStreamWriter({ xmlDeclaration: false });
    w.start('t');
    w.text('hello');
    w.end();
    expect(decode(w.result())).toBe('<t>hello</t>');
  });

  it('escapes attribute values and text', () => {
    const w = createXmlStreamWriter({ xmlDeclaration: false });
    w.start('a', { v: 'x"y\nz' });
    w.text('a & b');
    w.end();
    expect(decode(w.result())).toBe('<a v="x&quot;y&#10;z">a &amp; b</a>');
  });

  it('omits standalone when requested', () => {
    const w = createXmlStreamWriter({ standalone: 'omit' });
    w.start('r');
    w.end();
    expect(decode(w.result())).toBe('<?xml version="1.0" encoding="UTF-8"?>\n<r/>');
  });
});

describe('createXmlStreamWriter — namespacing', () => {
  it('prefixes Clark-named elements via DEFAULT_PREFIXES', () => {
    const w = createXmlStreamWriter({ xmlDeclaration: false });
    w.start(`{${SHEET_MAIN_NS}}sheet`, {
      xmlns: SHEET_MAIN_NS,
      [`xmlns:r`]: REL_NS,
      [`{${REL_NS}}id`]: 'rId1',
    });
    w.end();
    const xml = decode(w.result());
    expect(xml).toContain('<sheet ');
    expect(xml).toContain(`xmlns="${SHEET_MAIN_NS}"`);
    expect(xml).toContain(`xmlns:r="${REL_NS}"`);
    expect(xml).toContain('r:id="rId1"');
  });

  it('uses xml: as the canonical prefix for XML_NS attributes', () => {
    const w = createXmlStreamWriter({ xmlDeclaration: false });
    w.start('t', { [`{${XML_NS}}space`]: 'preserve' });
    w.text('  hi  ');
    w.end();
    expect(decode(w.result())).toBe('<t xml:space="preserve">  hi  </t>');
  });

  it('lets the caller override the prefix map', () => {
    const w = createXmlStreamWriter({
      xmlDeclaration: false,
      prefixMap: { [CHART_NS]: 'chart' },
    });
    w.start(`{${CHART_NS}}chart`);
    w.end();
    expect(decode(w.result())).toBe('<chart:chart/>');
  });
});

describe('createXmlStreamWriter — writeNode + writeRaw', () => {
  it('splices a complete subtree via writeNode', () => {
    const w = createXmlStreamWriter({ xmlDeclaration: false });
    w.start('root');
    w.writeNode(el('child', { a: '1' }, [el('leaf', {}, [], 'X')]));
    w.end();
    expect(decode(w.result())).toBe('<root><child a="1"><leaf>X</leaf></child></root>');
  });

  it('writeRaw appends bytes verbatim — used by the cell hot path', () => {
    const w = createXmlStreamWriter({ xmlDeclaration: false });
    w.start('sheetData');
    w.writeRaw('<row r="1"><c r="A1" t="s"><v>0</v></c></row>');
    w.end();
    expect(decode(w.result())).toBe('<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>');
  });
});

describe('createXmlStreamWriter — round-trip with parseXml', () => {
  it('produces XML that parseXml can re-read into the same logical tree', () => {
    const w = createXmlStreamWriter();
    w.start(`{${SHEET_MAIN_NS}}workbook`, {
      xmlns: SHEET_MAIN_NS,
      [`xmlns:r`]: REL_NS,
    });
    w.start(`{${SHEET_MAIN_NS}}sheets`);
    for (let i = 1; i <= 3; i++) {
      w.start(`{${SHEET_MAIN_NS}}sheet`, {
        name: `Sheet${i}`,
        sheetId: String(i),
        [`{${REL_NS}}id`]: `rId${i}`,
      });
      w.end();
    }
    w.end();
    w.end();

    const root = parseXml(w.result());
    expect(root.name).toBe(`{${SHEET_MAIN_NS}}workbook`);
    const sheets = root.children.find((c) => c.name === `{${SHEET_MAIN_NS}}sheets`);
    expect(sheets?.children.length).toBe(3);
  });
});

describe('createXmlStreamWriter — error paths', () => {
  it('end() with no open element throws', () => {
    const w = createXmlStreamWriter();
    expect(() => w.end()).toThrowError(OpenXmlIoError);
  });

  it('result() with unclosed elements throws', () => {
    const w = createXmlStreamWriter({ xmlDeclaration: false });
    w.start('r');
    expect(() => w.result()).toThrowError(OpenXmlIoError);
  });

  it('writes after result() throw', () => {
    const w = createXmlStreamWriter({ xmlDeclaration: false });
    w.start('r');
    w.end();
    w.result();
    expect(() => w.start('x')).toThrowError(OpenXmlIoError);
    expect(() => w.text('y')).toThrowError(OpenXmlIoError);
    expect(() => w.writeRaw('z')).toThrowError(OpenXmlIoError);
  });
});

describe('createXmlStreamWriter — bulk output sanity', () => {
  it('emits 100k <c> elements via writeRaw + start/end and parses back to a balanced tree', () => {
    const w = createXmlStreamWriter({ xmlDeclaration: false });
    w.start('sheetData');
    const N = 100_000;
    for (let i = 1; i <= N; i++) {
      w.writeRaw(`<c r="A${i}"><v>${i}</v></c>`);
    }
    w.end();

    const bytes = w.result();
    // Loose sanity: bytes should be well over the auto-flush threshold so
    // we know the flush path was exercised at least once.
    expect(bytes.byteLength).toBeGreaterThan(1024 * 1024);

    // Parse back via DOM parser; element count should match.
    const root = parseXml(bytes);
    expect(root.name).toBe('sheetData');
    expect(root.children.length).toBe(N);
  }, 15_000);
});
