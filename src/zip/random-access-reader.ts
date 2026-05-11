// Random-access ZIP reader. / §2.3 streaming-read residual.
//
// The previous reader handed every entry to `fflate.unzipSync` up front, which
// materialises *every* uncompressed payload into memory at once. For a 100 MB
// xlsx with ~500 MB of decompressed sheet data the resident set spikes
// accordingly. The random-access path keeps only the compressed archive bytes
// resident, parses the central directory once (cheap — ~46 B per entry plus
// filename), and inflates each entry lazily on `read(path)`.
//
// Limitations:
// - ZIP64 reads only when the standard ZIP32 fields fit. EOCD with
// sentinel values (0xFFFF / 0xFFFFFFFF) falls back to fflate's `unzipSync` so
// external ZIP64 archives still load (the writer side has its own ZIP32 cap
// guard).
// - Compression methods: STORE (0) and DEFLATE (8). Anything else
// throws OpenXmlIoError.

import { Inflate, inflateSync, unzipSync } from 'fflate';
import { OpenXmlIoError } from '../utils/exceptions';
import type { ZipArchive } from './reader';

/**
 * Chunk size used when feeding compressed bytes into fflate's `Inflate` for
 * streaming reads. 64 KB matches the saxes/SAX consumer's typical batch and
 * keeps peak transient memory bounded.
 */
const INFLATE_CHUNK_BYTES = 64 * 1024;

const singleChunkStream = (bytes: Uint8Array): ReadableStream<Uint8Array> =>
  new ReadableStream<Uint8Array>({
    start(controller) {
      if (bytes.byteLength > 0) controller.enqueue(bytes);
      controller.close();
    },
  });

const SIG_EOCD = 0x06054b50;
const SIG_CD = 0x02014b50;
const SIG_LFH = 0x04034b50;
const COMP_STORE = 0;
const COMP_DEFLATE = 8;

interface CdEntry {
  path: string;
  lfhOffset: number;
  compMethod: number;
  compSize: number;
  uncompSize: number;
  /** General-purpose bit flag, used to detect bit-3 (data descriptor) and bit-11 (UTF-8). */
  gpFlag: number;
}

const u16 = (b: Uint8Array, off: number): number => (b[off] ?? 0) | ((b[off + 1] ?? 0) << 8);
const u32 = (b: Uint8Array, off: number): number => {
  const v0 = b[off] ?? 0;
  const v1 = b[off + 1] ?? 0;
  const v2 = b[off + 2] ?? 0;
  const v3 = b[off + 3] ?? 0;
  return (v0 | (v1 << 8) | (v2 << 16) | (v3 << 24)) >>> 0;
};

/** Find the End-of-Central-Directory record by scanning backwards from EOF. */
function findEocd(b: Uint8Array): number {
  const minStart = Math.max(0, b.length - 22 - 0xffff);
  for (let i = b.length - 22; i >= minStart; i--) {
    if (u32(b, i) === SIG_EOCD) return i;
  }
  throw new OpenXmlIoError('openZip: no End-of-Central-Directory signature found');
}

/** Parse the central directory into an array of entry descriptors. */
function parseCentralDirectory(b: Uint8Array, cdOffset: number, expectedCount: number): CdEntry[] {
  const entries: CdEntry[] = [];
  let p = cdOffset;
  for (let i = 0; i < expectedCount; i++) {
    if (u32(b, p) !== SIG_CD) {
      throw new OpenXmlIoError(`openZip: malformed central directory at byte ${p}`);
    }
    const gpFlag = u16(b, p + 8);
    const compMethod = u16(b, p + 10);
    const compSize = u32(b, p + 20);
    const uncompSize = u32(b, p + 24);
    const nameLen = u16(b, p + 28);
    const extraLen = u16(b, p + 30);
    const commentLen = u16(b, p + 32);
    const lfhOffset = u32(b, p + 42);
    const nameBytes = b.subarray(p + 46, p + 46 + nameLen);
    // Bit 11 (0x0800) signals UTF-8 filename. xlsx archives are almost always
    // UTF-8 already; treat bit-0 as UTF-8 too since CP437 ⊃ ASCII and xlsx uses
    // ASCII paths.
    const path = new TextDecoder('utf-8').decode(nameBytes);
    entries.push({ path, lfhOffset, compMethod, compSize, uncompSize, gpFlag });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

/** Read the compressed bytes for a CD entry by walking its local file header. */
function readCompressedBytes(b: Uint8Array, entry: CdEntry): Uint8Array {
  if (u32(b, entry.lfhOffset) !== SIG_LFH) {
    throw new OpenXmlIoError(`openZip: malformed local file header for "${entry.path}"`);
  }
  const nameLen = u16(b, entry.lfhOffset + 26);
  const extraLen = u16(b, entry.lfhOffset + 28);
  const dataStart = entry.lfhOffset + 30 + nameLen + extraLen;
  return b.subarray(dataStart, dataStart + entry.compSize);
}

/**
 * Open a buffered xlsx archive in random-access mode. The archive bytes stay
 * resident; entries inflate on demand inside `read(path)`.
 *
 * Falls back to `fflate.unzipSync` when the central directory uses ZIP64
 * sentinel values (entry count == 0xFFFF or any size field == 0xFFFFFFFF) so
 * external ZIP64 archives still load. xlsx files in the wild fit comfortably in
 * ZIP32; the fallback exists for safety.
 */
export function openRandomAccessArchive(bytes: Uint8Array): ZipArchive {
  // Quick sanity on min archive size.
  if (bytes.length < 22) {
    throw new OpenXmlIoError('openZip: archive is shorter than the minimum EOCD size (22 bytes)');
  }

  let eocdOff: number;
  try {
    eocdOff = findEocd(bytes);
  } catch (cause) {
    throw new OpenXmlIoError('openZip: archive is not a valid zip', { cause });
  }

  const totalEntries = u16(bytes, eocdOff + 10);
  const cdSize = u32(bytes, eocdOff + 12);
  const cdOffset = u32(bytes, eocdOff + 16);

  // ZIP64 fallback — fflate's unzipSync handles the extended record.
  if (totalEntries === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff) {
    return openViaUnzipSync(bytes);
  }

  let entries: CdEntry[];
  try {
    entries = parseCentralDirectory(bytes, cdOffset, totalEntries);
  } catch {
    // Malformed CD — fall back to fflate which is more tolerant.
    return openViaUnzipSync(bytes);
  }

  const byPath = new Map<string, CdEntry>();
  for (const e of entries) byPath.set(e.path, e);

  // Per-entry inflate cache so repeated reads of the same path don't re-inflate
  // — `read(path)` is documented as cheap on the second call (loadWorkbook
  // touches several files multiple times).
  const inflateCache = new Map<string, Uint8Array>();
  let live = true;
  let archiveBytes: Uint8Array | undefined = bytes;

  const ensureLive = (): Uint8Array => {
    if (!live || !archiveBytes) {
      throw new OpenXmlIoError('openZip: archive is closed');
    }
    return archiveBytes;
  };

  const readEntry = (path: string): Uint8Array => {
    const buf = ensureLive();
    const cached = inflateCache.get(path);
    if (cached) return cached;
    const entry = byPath.get(path);
    if (!entry) {
      throw new OpenXmlIoError(`openZip: no entry at "${path}"`);
    }
    const compressed = readCompressedBytes(buf, entry);
    let out: Uint8Array;
    if (entry.compMethod === COMP_STORE) {
      // Copy so callers can safely mutate the returned bytes without perturbing
      // the underlying archive view.
      out = compressed.slice();
    } else if (entry.compMethod === COMP_DEFLATE) {
      try {
        out = inflateSync(compressed);
      } catch (cause) {
        throw new OpenXmlIoError(`openZip: failed to inflate "${path}"`, { cause });
      }
    } else {
      throw new OpenXmlIoError(`openZip: unsupported compression method ${entry.compMethod} for "${path}"`);
    }
    inflateCache.set(path, out);
    return out;
  };

  const readEntryStream = (path: string): ReadableStream<Uint8Array> => {
    const buf = ensureLive();
    const cached = inflateCache.get(path);
    if (cached) return singleChunkStream(cached);
    const entry = byPath.get(path);
    if (!entry) {
      throw new OpenXmlIoError(`openZip: no entry at "${path}"`);
    }
    const compressed = readCompressedBytes(buf, entry);
    if (entry.compMethod === COMP_STORE) {
      // STORE means the bytes on disk are the bytes the caller wants; no need
      // to involve the inflate state machine. Hand them out as a single chunk
      // — copy because callers may mutate the returned bytes.
      return singleChunkStream(compressed.slice());
    }
    if (entry.compMethod !== COMP_DEFLATE) {
      throw new OpenXmlIoError(`openZip: unsupported compression method ${entry.compMethod} for "${path}"`);
    }
    // DEFLATE: feed the compressed bytes into fflate's `Inflate` in fixed-size
    // chunks and forward each inflated chunk to the stream consumer. Peak
    // memory stays at the chunk size + the inflate state, never the full
    // uncompressed payload.
    return new ReadableStream<Uint8Array>({
      start(controller) {
        let errored = false;
        const inflater = new Inflate((chunk, final) => {
          if (errored) return;
          if (chunk.byteLength > 0) controller.enqueue(chunk);
          if (final) controller.close();
        });
        try {
          let off = 0;
          while (off < compressed.byteLength) {
            const end = Math.min(off + INFLATE_CHUNK_BYTES, compressed.byteLength);
            const slice = compressed.subarray(off, end);
            const isLast = end >= compressed.byteLength;
            inflater.push(slice, isLast);
            off = end;
          }
        } catch (cause) {
          errored = true;
          controller.error(new OpenXmlIoError(`openZip: failed to inflate "${path}"`, { cause }));
        }
      },
    });
  };

  return {
    list(): string[] {
      ensureLive();
      return [...byPath.keys()].sort();
    },
    has(path: string): boolean {
      if (!live) return false;
      return byPath.has(path);
    },
    read(path: string): Uint8Array {
      return readEntry(path);
    },
    async readAsync(path: string): Promise<Uint8Array> {
      return readEntry(path);
    },
    readStream(path: string): ReadableStream<Uint8Array> {
      return readEntryStream(path);
    },
    close(): void {
      live = false;
      archiveBytes = undefined;
      inflateCache.clear();
      byPath.clear();
    },
  };
}

/** Fallback: hand the whole archive to fflate when ZIP64 / unsupported features turn up. */
function openViaUnzipSync(bytes: Uint8Array): ZipArchive {
  let entries: Record<string, Uint8Array> | undefined;
  try {
    entries = unzipSync(bytes);
  } catch (cause) {
    throw new OpenXmlIoError('openZip: archive is not a valid zip', { cause });
  }
  let live = true;
  return {
    list(): string[] {
      if (!live || !entries) throw new OpenXmlIoError('openZip: archive is closed');
      return Object.keys(entries).sort();
    },
    has(path: string): boolean {
      if (!live || !entries) return false;
      return Object.hasOwn(entries, path);
    },
    read(path: string): Uint8Array {
      if (!live || !entries) throw new OpenXmlIoError('openZip: archive is closed');
      const e = entries[path];
      if (!e) throw new OpenXmlIoError(`openZip: no entry at "${path}"`);
      return e;
    },
    async readAsync(path: string): Promise<Uint8Array> {
      return this.read(path);
    },
    readStream(path: string): ReadableStream<Uint8Array> {
      // The unzipSync fallback path already has the entry fully inflated
      // (that's the price of dropping back from random-access). Hand it out as
      // a single-chunk stream so callers using the streaming reader don't have
      // to branch on the implementation.
      const inflated = this.read(path);
      return singleChunkStream(inflated);
    },
    close(): void {
      live = false;
      entries = undefined;
    },
  };
}
