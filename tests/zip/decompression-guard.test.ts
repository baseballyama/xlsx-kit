// Decompression-bomb guard behaviour. Confirms the limits described on
// `DecompressionLimits` actually fire on adversarial archives and that
// legitimate archives still load. Each scenario exercises one independent
// bound: per-entry size, per-archive total, and compression ratio.

import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node';
import { OpenXmlDecompressionBombError } from '../../src/utils/exceptions';
import { DEFAULT_DECOMPRESSION_LIMITS } from '../../src/zip/decompression-guard';
import { openZip } from '../../src/zip/reader';
import { createZipWriter } from '../../src/zip/writer';

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

// A run of identical bytes deflates to a tiny payload, producing extreme
// compression ratios that simulate a zip-bomb without actually allocating
// gigabytes.
const zeros = (n: number): Uint8Array => new Uint8Array(n);

describe('decompression-bomb guard — defaults', () => {
  it('admits legitimately compressible data well below the cap', async () => {
    // 1 MiB of zeros deflates to ~1 KiB → ratio ~1024 but tiny absolute size.
    // The default 1000× ceiling means a 1 KiB compressed entry can reach
    // ~1 MiB uncompressed, which this just barely sits on; the 64 B
    // `RATIO_CHECK_MIN_COMPRESSED_BYTES` floor isn't reached here, but the
    // entry / total bounds give it room.
    const payload = new TextEncoder().encode('hello world '.repeat(1000));
    const bytes = await buildArchive([{ path: 'a.txt', bytes: payload }]);
    const archive = await openZip(fromBuffer(bytes));
    expect(archive.read('a.txt')).toEqual(payload);
    archive.close();
  });

  it('rejects an entry whose central-directory ratio is above the ceiling', async () => {
    // 8 MiB of zeros deflates to a couple KiB. Ratio is ~4000× — well above
    // the default 1000× and a textbook zip-bomb signature.
    const bytes = await buildArchive([{ path: 'bomb.bin', bytes: zeros(8 * 1024 * 1024) }]);
    await expect(openZip(fromBuffer(bytes))).rejects.toThrowError(OpenXmlDecompressionBombError);
  });

  it('rejects an entry whose central-directory uncompressed size exceeds the per-entry cap', async () => {
    // Forge an archive that lies in the CD: declare a uncompSize beyond the
    // limit. We can't easily produce a >512 MiB legitimate compressed entry in
    // a test, so simulate via a custom limit instead — the actual default cap
    // is exercised here by the smaller `maxEntryUncompressedBytes` override.
    const bytes = await buildArchive([{ path: 'big.bin', bytes: zeros(2 * 1024 * 1024) }]);
    await expect(
      openZip(fromBuffer(bytes), { decompressionLimits: { maxEntryUncompressedBytes: 1024 } }),
    ).rejects.toThrowError(/per-entry limit/);
  });

  it('rejects when the summed declared uncompressed bytes exceed the archive cap', async () => {
    // Three 1 MiB entries with a 2 MiB archive cap — the third declared total
    // crosses the limit. Use a generous ratio override so the ratio check on
    // each individual entry (~1014×) doesn't fire before the total check.
    const bytes = await buildArchive([
      { path: 'a.bin', bytes: zeros(1024 * 1024) },
      { path: 'b.bin', bytes: zeros(1024 * 1024) },
      { path: 'c.bin', bytes: zeros(1024 * 1024) },
    ]);
    await expect(
      openZip(fromBuffer(bytes), {
        decompressionLimits: {
          maxTotalUncompressedBytes: 2 * 1024 * 1024,
          maxCompressionRatio: 100_000,
        },
      }),
    ).rejects.toThrowError(/archive limit/);
  });

  it('reports the bomb via OpenXmlDecompressionBombError (subclass of OpenXmlIoError)', async () => {
    const bytes = await buildArchive([{ path: 'bomb.bin', bytes: zeros(8 * 1024 * 1024) }]);
    try {
      await openZip(fromBuffer(bytes));
      throw new Error('expected to throw');
    } catch (err) {
      expect(err).toBeInstanceOf(OpenXmlDecompressionBombError);
      // Subclass of OpenXmlIoError so existing catch-all paths keep working.
      expect((err as Error).message).toMatch(/decompression-bomb/);
    }
  });
});

describe('decompression-bomb guard — overrides', () => {
  it('decompressionLimits: false disables every check', async () => {
    const bytes = await buildArchive([{ path: 'bomb.bin', bytes: zeros(8 * 1024 * 1024) }]);
    const archive = await openZip(fromBuffer(bytes), { decompressionLimits: false });
    const out = archive.read('bomb.bin');
    expect(out.byteLength).toBe(8 * 1024 * 1024);
    archive.close();
  });

  it('partial override merges with defaults', async () => {
    // Only override the ratio; the per-entry / total caps stay at the defaults.
    // 8 MiB / ~2 KiB ratio sails under a 10000× ceiling.
    const bytes = await buildArchive([{ path: 'soft.bin', bytes: zeros(8 * 1024 * 1024) }]);
    const archive = await openZip(fromBuffer(bytes), {
      decompressionLimits: { maxCompressionRatio: 10_000 },
    });
    expect(archive.read('soft.bin').byteLength).toBe(8 * 1024 * 1024);
    archive.close();
  });

  it('the DEFAULT_DECOMPRESSION_LIMITS constant is the documented baseline', () => {
    expect(DEFAULT_DECOMPRESSION_LIMITS.maxEntryUncompressedBytes).toBe(512 * 1024 * 1024);
    expect(DEFAULT_DECOMPRESSION_LIMITS.maxTotalUncompressedBytes).toBe(1024 * 1024 * 1024);
    expect(DEFAULT_DECOMPRESSION_LIMITS.maxCompressionRatio).toBe(1000);
  });
});

// Patch the uncompSize fields (Central Directory + Local File Header) of the
// only entry in `bytes` so the CD declares a smaller size than the entry
// actually inflates to. This is how a real bomb hides its true size from a
// CD-pre-check — needed to exercise the runtime inflate-time abort.
const patchSingleEntryUncompSize = (bytes: Uint8Array, declaredUncomp: number): Uint8Array => {
  const SIG_CD = 0x02014b50;
  const SIG_LFH = 0x04034b50;
  const u32 = (b: Uint8Array, off: number): number =>
    ((b[off] ?? 0) | ((b[off + 1] ?? 0) << 8) | ((b[off + 2] ?? 0) << 16) | ((b[off + 3] ?? 0) << 24)) >>> 0;
  const writeU32 = (b: Uint8Array, off: number, v: number): void => {
    b[off] = v & 0xff;
    b[off + 1] = (v >>> 8) & 0xff;
    b[off + 2] = (v >>> 16) & 0xff;
    b[off + 3] = (v >>> 24) & 0xff;
  };
  const out = new Uint8Array(bytes);
  // Find LFH (single entry: starts at offset 0).
  if (u32(out, 0) !== SIG_LFH) throw new Error('expected LFH at offset 0');
  writeU32(out, 22, declaredUncomp); // LFH +22 == uncompressed size
  // Find CD (scan forward from LFH end).
  for (let i = 0; i < out.length - 4; i++) {
    if (u32(out, i) === SIG_CD) {
      writeU32(out, i + 24, declaredUncomp); // CD +24 == uncompressed size
      return out;
    }
  }
  throw new Error('no CD found');
};

describe('decompression-bomb guard — streaming reads', () => {
  it('aborts a streaming read when the inflated size crosses the runtime cap despite an honest-looking CD', async () => {
    // Build a 4 MiB-zero entry, then mutate the CD so it claims only 1 KiB.
    // The pre-check passes (declared sizes are tiny), but inflate produces the
    // real payload — the streaming abort must fire mid-flight.
    const honest = await buildArchive([{ path: 'big.bin', bytes: zeros(4 * 1024 * 1024) }]);
    const lying = patchSingleEntryUncompSize(honest, 1024);
    const archive = await openZip(fromBuffer(lying), {
      decompressionLimits: { maxEntryUncompressedBytes: 64 * 1024, maxCompressionRatio: 1_000_000 },
    });
    const stream = archive.readStream('big.bin');
    const reader = stream.getReader();
    await expect(
      (async () => {
        while (true) {
          const { done } = await reader.read();
          if (done) break;
        }
      })(),
    ).rejects.toThrowError(/decompression-bomb/);
    archive.close();
  });

  it('aborts a sync read on a CD-lying entry too', async () => {
    const honest = await buildArchive([{ path: 'big.bin', bytes: zeros(4 * 1024 * 1024) }]);
    const lying = patchSingleEntryUncompSize(honest, 1024);
    const archive = await openZip(fromBuffer(lying), {
      decompressionLimits: { maxEntryUncompressedBytes: 64 * 1024, maxCompressionRatio: 1_000_000 },
    });
    expect(() => archive.read('big.bin')).toThrowError(/decompression-bomb/);
    archive.close();
  });
});
