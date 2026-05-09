// Random-access ZIP reader correctness. Verifies that openZip walks the central
// directory once and inflates entries on demand — reading in any order,
// repeating reads, and mixing STORE / DEFLATE entries all land on
// byte-identical payloads.

import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../../src/io/node';
import { openZip } from '../../../src/zip/reader';
import { createZipWriter } from '../../../src/zip/writer';

const buildArchive = async (
  entries: ReadonlyArray<{ path: string; bytes: Uint8Array; compress?: boolean }>,
): Promise<Uint8Array> => {
  const sink = toBuffer();
  const w = createZipWriter(sink);
  for (const e of entries) {
    await w.addEntry(e.path, e.bytes, e.compress === false ? { compress: false } : undefined);
  }
  await w.finalize();
  return sink.result();
};

describe('openZip — random-access reader', () => {
  it('reads entries out of central-directory order without disturbing other entries', async () => {
    const a = new Uint8Array([1, 2, 3]);
    const b = new TextEncoder().encode('a string with enough text for deflate to actually run');
    const c = new Uint8Array([0xff, 0xee, 0xdd, 0xcc]);
    const bytes = await buildArchive([
      { path: 'a.bin', bytes: a, compress: false },
      { path: 'b.txt', bytes: b },
      { path: 'c.bin', bytes: c, compress: false },
    ]);
    const archive = await openZip(fromBuffer(bytes));

    // Read tail-first to make sure offsets aren't computed from a sequential
    // cursor.
    expect(archive.read('c.bin')).toEqual(c);
    expect(archive.read('a.bin')).toEqual(a);
    expect(archive.read('b.txt')).toEqual(b);
    archive.close();
  });

  it('returns the same bytes on repeated reads (inflate cache)', async () => {
    const payload = new TextEncoder().encode('repeat me '.repeat(50));
    const bytes = await buildArchive([{ path: 'r.txt', bytes: payload }]);
    const archive = await openZip(fromBuffer(bytes));
    const first = archive.read('r.txt');
    const second = archive.read('r.txt');
    expect(first).toEqual(payload);
    expect(second).toEqual(first);
    archive.close();
  });

  it('rejects unknown paths with OpenXmlIoError', async () => {
    const bytes = await buildArchive([{ path: 'a.txt', bytes: new Uint8Array([1]) }]);
    const archive = await openZip(fromBuffer(bytes));
    expect(() => archive.read('nope.txt')).toThrowError('no entry at "nope.txt"');
    archive.close();
  });

  it('list() returns entries in lexical order', async () => {
    const bytes = await buildArchive([
      { path: 'z.txt', bytes: new Uint8Array([1]) },
      { path: 'a.txt', bytes: new Uint8Array([2]) },
      { path: 'm.txt', bytes: new Uint8Array([3]) },
    ]);
    const archive = await openZip(fromBuffer(bytes));
    expect(archive.list()).toEqual(['a.txt', 'm.txt', 'z.txt']);
    archive.close();
  });

  it('reads and writes after close throw OpenXmlIoError', async () => {
    const bytes = await buildArchive([{ path: 'a.txt', bytes: new Uint8Array([1]) }]);
    const archive = await openZip(fromBuffer(bytes));
    archive.close();
    expect(() => archive.read('a.txt')).toThrowError('archive is closed');
    expect(() => archive.list()).toThrowError('archive is closed');
    expect(archive.has('a.txt')).toBe(false);
  });

  it('rejects archives with no EOCD signature', async () => {
    const bogus = new Uint8Array(64); // zeroes — no EOCD anywhere
    await expect(openZip(fromBuffer(bogus))).rejects.toThrowError(/not a valid zip|too short/i);
  });
});
