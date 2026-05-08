// ZIP64 post-processing for archives whose entry count exceeds the
// 16-bit ZIP32 cap (65535). fflate's `Zip` writer always emits a
// plain ZIP32 EOCD; for archives with more entries we keep its
// per-entry LFH/CDH layout (correct as long as no individual size or
// offset overflows 32 bits) and splice in a ZIP64 End-of-Central-
// Directory record + locator before the EOCD, then patch the EOCD's
// entry-count fields with the 0xFFFF sentinel that signals "consult
// the ZIP64 record for the real values".
//
// The input is fflate's *final* chunk — the trailing [CD | EOCD] block
// it emits in one ondata callback when `Zip.end()` is called. The
// preceding entry-data chunks have already been streamed to the sink,
// so we operate on the final chunk in isolation. The global EOCD
// offset (needed for the ZIP64 locator) is derivable from cd_offset +
// cd_size carried in the EOCD itself, so no external bookkeeping is
// required.
//
// Out of scope: per-entry sizes or central-directory offsets > 4 GiB.
// xlsx archives don't approach those limits in practice; we throw a
// clear error if we detect overflow there.

import { OpenXmlIoError, OpenXmlNotImplementedError } from '../utils/exceptions';

const ZIP32_MAX_ENTRIES = 0xffff;
const ZIP32_MAX_U32 = 0xffffffff;
const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_EOCD_LOCATOR = 0x07064b50;

const ZIP64_EOCD_SIZE = 56;
const ZIP64_LOCATOR_SIZE = 20;

const u16 = (b: Uint8Array, o: number): number => (b[o] ?? 0) | ((b[o + 1] ?? 0) << 8);

const u32 = (b: Uint8Array, o: number): number => {
  const v0 = b[o] ?? 0;
  const v1 = b[o + 1] ?? 0;
  const v2 = b[o + 2] ?? 0;
  const v3 = b[o + 3] ?? 0;
  return (v0 | (v1 << 8) | (v2 << 16) | (v3 << 24)) >>> 0;
};

const writeU16 = (b: Uint8Array, o: number, v: number): void => {
  b[o] = v & 0xff;
  b[o + 1] = (v >>> 8) & 0xff;
};

const writeU32 = (b: Uint8Array, o: number, v: number): void => {
  b[o] = v & 0xff;
  b[o + 1] = (v >>> 8) & 0xff;
  b[o + 2] = (v >>> 16) & 0xff;
  b[o + 3] = (v >>> 24) & 0xff;
};

const writeU64 = (b: Uint8Array, o: number, v: number): void => {
  // JS Number safely represents integers up to 2^53 - 1, well beyond
  // anything we'd ever emit here. Split via Math.floor + modulo to
  // avoid bit-shift truncation at 32 bits.
  const lo = v >>> 0;
  const hi = Math.floor(v / 0x100000000) >>> 0;
  writeU32(b, o, lo);
  writeU32(b, o + 4, hi);
};

const findEocdOffset = (bytes: Uint8Array): number => {
  // EOCD is min 22 bytes and may be followed by up to 65535 bytes of
  // archive comment. Scan backwards from the latest possible position.
  const minOffset = Math.max(0, bytes.length - (22 + 0xffff));
  for (let p = bytes.length - 22; p >= minOffset; p--) {
    if (u32(bytes, p) === SIG_EOCD) {
      const commentLen = u16(bytes, p + 20);
      if (p + 22 + commentLen === bytes.length) return p;
    }
  }
  throw new OpenXmlIoError('zip64-patch: no End-of-Central-Directory signature found');
};

/**
 * Splice ZIP64 EOCD record + locator into fflate's final chunk and
 * patch the trailing EOCD entry-count fields with the 0xFFFF sentinel.
 *
 * `finalChunk` must be the [CD | EOCD] block fflate emits as its last
 * ondata callback (everything before it is per-entry LFH/data/DD that
 * we leave untouched). Returns a new chunk; the input is not mutated.
 *
 * Assumes per-entry sizes and central-directory offset fit in 32 bits;
 * throws if not (xlsx archives never approach those limits).
 */
export function applyZip64EntryCountPatch(finalChunk: Uint8Array, totalEntries: number): Uint8Array {
  if (totalEntries <= ZIP32_MAX_ENTRIES) return finalChunk;

  const eocdOffset = findEocdOffset(finalChunk);

  const cdSize = u32(finalChunk, eocdOffset + 12);
  const cdOffset = u32(finalChunk, eocdOffset + 16);
  const commentLen = u16(finalChunk, eocdOffset + 20);

  if (cdSize === ZIP32_MAX_U32 || cdOffset === ZIP32_MAX_U32) {
    throw new OpenXmlNotImplementedError(
      'zip64-patch: archive size or central-directory offset exceeds 4 GiB; full ZIP64 size support is not implemented (xlsx in practice stays well under 4 GiB).',
    );
  }

  // Where the EOCD starts in the *global* archive (before our patch).
  // CD precedes EOCD with no gap, so global EOCD offset is just
  // cd_offset + cd_size — the locator points here.
  const globalEocdOffset = cdOffset + cdSize;

  const eocdLen = 22 + commentLen;
  const newChunkLen = eocdOffset + ZIP64_EOCD_SIZE + ZIP64_LOCATOR_SIZE + eocdLen;
  const out = new Uint8Array(newChunkLen);

  // Original CD bytes (everything before the EOCD).
  out.set(finalChunk.subarray(0, eocdOffset), 0);

  // ZIP64 EOCD record (56 bytes total).
  const zip64Eocd = out.subarray(eocdOffset, eocdOffset + ZIP64_EOCD_SIZE);
  writeU32(zip64Eocd, 0, SIG_ZIP64_EOCD);
  // size_of_zip64_eocd = total_size - 12 (size field excludes signature + this field itself)
  writeU64(zip64Eocd, 4, ZIP64_EOCD_SIZE - 12);
  writeU16(zip64Eocd, 12, 45); // version made by (4.5 — first ZIP64 spec)
  writeU16(zip64Eocd, 14, 45); // version needed
  writeU32(zip64Eocd, 16, 0); // disk_number
  writeU32(zip64Eocd, 20, 0); // disk_with_cd
  writeU64(zip64Eocd, 24, totalEntries); // entries_on_this_disk
  writeU64(zip64Eocd, 32, totalEntries); // total_entries
  writeU64(zip64Eocd, 40, cdSize); // cd_size
  writeU64(zip64Eocd, 48, cdOffset); // cd_offset

  // ZIP64 EOCD locator (20 bytes).
  const locOffset = eocdOffset + ZIP64_EOCD_SIZE;
  const locator = out.subarray(locOffset, locOffset + ZIP64_LOCATOR_SIZE);
  writeU32(locator, 0, SIG_ZIP64_EOCD_LOCATOR);
  writeU32(locator, 4, 0); // disk_with_zip64_eocd
  writeU64(locator, 8, globalEocdOffset);
  writeU32(locator, 16, 1); // total_disks

  // New EOCD: copy original then patch entry counts to the 0xFFFF
  // sentinel. (fflate writes the low 16 bits of the true count there,
  // which confuses readers that don't first look for the ZIP64 record.)
  const newEocdOffset = locOffset + ZIP64_LOCATOR_SIZE;
  out.set(finalChunk.subarray(eocdOffset, eocdOffset + eocdLen), newEocdOffset);
  writeU16(out, newEocdOffset + 8, ZIP32_MAX_ENTRIES);
  writeU16(out, newEocdOffset + 10, ZIP32_MAX_ENTRIES);

  return out;
}
