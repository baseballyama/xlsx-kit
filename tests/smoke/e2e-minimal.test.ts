// Phase-1 §10 end-to-end smoke test:
//
//   "openpyxl が作った最小 xlsx を解凍 → manifest を読む → 元と等価な
//    manifest を出力 → 再ジップ"
//
// Touches every layer the bootstrap blocks delivered:
// - io   (fromBuffer, toBuffer)
// - zip  (openZip, createZipWriter)
// - xml  (parseXml, serializeXml; via packaging)
// - schema (toTree / fromTree; via packaging)
// - packaging (manifest, relationships, core / extended properties)
// - utils (exception types)

import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node';
import { corePropsFromBytes } from '../../src/packaging/core';
import { extendedPropsFromBytes } from '../../src/packaging/extended';
import { findOverride, manifestFromBytes, manifestToBytes } from '../../src/packaging/manifest';
import { findByType, relsFromBytes, relsToBytes } from '../../src/packaging/relationships';
import { openZip } from '../../src/zip/reader';
import { createZipWriter } from '../../src/zip/writer';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

const T_OFFICE_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

describe('phase-1 e2e: openpyxl genuine/empty.xlsx', () => {
  it('round-trips the manifest and root rels through the full stack', async () => {
    // ---- read --------------------------------------------------------------
    const original = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const manifest = manifestFromBytes(original.read('[Content_Types].xml'));
    const rootRels = relsFromBytes(original.read('_rels/.rels'));
    const coreProps = corePropsFromBytes(original.read('docProps/core.xml'));
    const appProps = extendedPropsFromBytes(original.read('docProps/app.xml'));

    // Sanity: structural facts about the fixture.
    expect(findOverride(manifest, '/xl/workbook.xml')?.contentType).toContain('spreadsheetml.sheet.main');
    expect(findByType(rootRels, T_OFFICE_DOC)?.target.endsWith('workbook.xml')).toBe(true);
    expect(coreProps.creator).toBe('CED501');
    expect(appProps.application).toBe('Microsoft Excel');

    // ---- write -------------------------------------------------------------
    const sink = toBuffer();
    const writer = createZipWriter(sink);

    // Re-emit the manifest and root rels through our own serialisers.
    await writer.addEntry('[Content_Types].xml', manifestToBytes(manifest));
    await writer.addEntry('_rels/.rels', relsToBytes(rootRels));

    // Pass the rest through verbatim — phase-1 doesn't model workbook
    // contents yet; that's phase 3.
    for (const path of original.list()) {
      if (path === '[Content_Types].xml' || path === '_rels/.rels') continue;
      const compress = !(path.startsWith('xl/media/') || path === 'xl/vbaProject.bin');
      await writer.addEntry(path, original.read(path), { compress });
    }
    await writer.finalize();

    // ---- re-read and verify ------------------------------------------------
    const round = await openZip(fromBuffer(sink.result()));
    expect(round.list().sort()).toEqual(original.list().sort());

    const m2 = manifestFromBytes(round.read('[Content_Types].xml'));
    const r2 = relsFromBytes(round.read('_rels/.rels'));
    expect(m2).toEqual(manifest);
    expect(r2).toEqual(rootRels);

    // The non-rebuilt parts must be preserved byte-for-byte.
    for (const path of round.list()) {
      if (path === '[Content_Types].xml' || path === '_rels/.rels') continue;
      expect(round.read(path)).toEqual(original.read(path));
    }
  });
});
