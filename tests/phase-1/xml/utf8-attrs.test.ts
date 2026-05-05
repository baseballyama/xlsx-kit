// Phase 1 §3 — UTF-8 attribute / element round-trip parity.
// Real-world xlsx files often carry non-ASCII text (sheet titles,
// codeName, content, comments) and the parser/serializer must
// preserve it byte-identical through the SAX iter and DOM walk.

import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { parseXml } from '../../../src/xml/parser';
import { serializeXml } from '../../../src/xml/serializer';

describe('phase-1 §3 — UTF-8 round-trip via parseXml + serializeXml', () => {
  it('preserves a Cyrillic workbookPr@codeName ("ЭтаКнига") through DOM', () => {
    const path = resolve(
      __dirname,
      '../../../reference/openpyxl/openpyxl/packaging/tests/data/workbook_russian_code_name.xml',
    );
    const bytes = readFileSync(path);
    const root = parseXml(bytes);

    // Walk to <workbookPr> and confirm codeName survives the parse.
    const findChild = (n: typeof root, local: string): typeof root | undefined => {
      for (const c of n.children) {
        if (c.name.endsWith(`}${local}`) || c.name === local) return c;
      }
      return undefined;
    };
    const pr = findChild(root, 'workbookPr');
    expect(pr).toBeDefined();
    expect(pr?.attrs['codeName']).toBe('ЭтаКнига');

    // Serialize back to bytes and re-parse — the codeName must round-trip.
    const out = serializeXml(root);
    const root2 = parseXml(out);
    const pr2 = findChild(root2, 'workbookPr');
    expect(pr2?.attrs['codeName']).toBe('ЭтаКнига');
  });

  it('preserves multi-byte text nodes (Japanese / Arabic / emoji)', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<root xmlns="urn:t"><a>こんにちは, ${String.fromCodePoint(0x1f600)} مرحبا</a><b attr="日本語"/></root>`;
    const root = parseXml(new TextEncoder().encode(xml));
    expect(root.children[0]?.text).toContain('こんにちは');
    expect(root.children[0]?.text).toContain('\u{1F600}');
    expect(root.children[0]?.text).toContain('مرحبا');
    expect(root.children[1]?.attrs['attr']).toBe('日本語');

    const back = parseXml(serializeXml(root));
    expect(back.children[0]?.text).toBe(root.children[0]?.text);
    expect(back.children[1]?.attrs['attr']).toBe('日本語');
  });
});
