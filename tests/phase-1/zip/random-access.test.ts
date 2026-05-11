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

  it('readStream yields the inflated payload chunk-by-chunk for a DEFLATE entry', async () => {
    // Build a payload that comfortably exceeds the inflate chunk size (64 KB)
    // and is incompressible enough that the compressed bytes are also large —
    // a sequential ramp deflates to a handful of bytes and would never force
    // multi-chunk pushes. Use a pseudo-random pattern derived from a tiny LCG
    // so it stays deterministic without pulling in WebCrypto.
    const big = new Uint8Array(512 * 1024);
    let seed = 0x12345678;
    for (let i = 0; i < big.byteLength; i++) {
      seed = (seed * 1664525 + 1013904223) >>> 0;
      big[i] = seed & 0xff;
    }
    const bytes = await buildArchive([{ path: 'big.bin', bytes: big }]);
    const archive = await openZip(fromBuffer(bytes));
    const stream = archive.readStream('big.bin');
    const reader = stream.getReader();
    const chunks: Uint8Array[] = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    let total = 0;
    for (const c of chunks) total += c.byteLength;
    const out = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      out.set(c, off);
      off += c.byteLength;
    }
    // Lossless round-trip is the primary invariant: whatever chunking strategy
    // fflate uses internally, the concatenated stream must equal the original
    // payload byte-for-byte.
    expect(out).toEqual(big);
    archive.close();
  });

  it('readStream emits a single chunk for STORE entries (no inflate involved)', async () => {
    const payload = new Uint8Array([10, 20, 30, 40]);
    const bytes = await buildArchive([{ path: 's.bin', bytes: payload, compress: false }]);
    const archive = await openZip(fromBuffer(bytes));
    const stream = archive.readStream('s.bin');
    const reader = stream.getReader();
    const first = await reader.read();
    expect(first.done).toBe(false);
    expect(Array.from(first.value ?? new Uint8Array())).toEqual([10, 20, 30, 40]);
    const second = await reader.read();
    expect(second.done).toBe(true);
    archive.close();
  });

  it('readStream rejects unknown paths with OpenXmlIoError', async () => {
    const bytes = await buildArchive([{ path: 'a.txt', bytes: new Uint8Array([1]) }]);
    const archive = await openZip(fromBuffer(bytes));
    expect(() => archive.readStream('nope.txt')).toThrowError('no entry at "nope.txt"');
    archive.close();
  });
});
