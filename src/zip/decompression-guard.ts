// Decompression-bomb safeguards for the ZIP reader.
//
// Zip-bombs encode pathological compression ratios (e.g. 100 B → 1 GB) so that
// a naive `inflate` exhausts memory before the caller can react. The guards
// here apply three orthogonal bounds:
//
// 1. Per-entry uncompressed cap — hard ceiling on a single payload.
// 2. Total-archive uncompressed cap — hard ceiling summed across the archive.
// 3. Per-entry compression ratio — bounds the *amplification factor* so that
//    even small compressed entries can't decompress into hundreds of MB.
//
// Defaults are sized to admit any legitimate xlsx (xml compresses well but
// rarely exceeds the per-entry cap; office templates with embedded media stay
// well under the total cap) while rejecting plausible bombs. Trusted callers
// (e.g. a backend that only ever loads xlsx it generated itself) can pass
// `false` to disable the guard.

import { OpenXmlDecompressionBombError } from '../utils/exceptions';

/** Per-archive limits enforced during {@link openZip}. */
export interface DecompressionLimits {
  /**
   * Maximum decompressed bytes for a single archive entry. Default 512 MiB.
   * Legitimate xlsx sheets stay well below this even with millions of cells.
   */
  maxEntryUncompressedBytes?: number;
  /**
   * Maximum decompressed bytes summed across every entry the caller reads.
   * Default 1 GiB.
   */
  maxTotalUncompressedBytes?: number;
  /**
   * Maximum allowed `uncompressed / compressed` ratio for a single entry.
   * Default 1000. xml usually compresses 5–20×, but highly repetitive payloads
   * (long runs of zeros, sparse worksheets) can hit several hundred ×
   * legitimately. Classic zip-bombs run 10 000× and up, so a 1000× ceiling
   * still catches them while admitting realistic content. The implementation
   * treats compressed sizes below 64 B as exempt — there's no amplification
   * budget worth policing on such small entries, and the absolute per-entry /
   * archive limits already cover them.
   */
  maxCompressionRatio?: number;
}

/** Resolved limits with no `undefined` fields. */
export interface ResolvedDecompressionLimits {
  readonly maxEntryUncompressedBytes: number;
  readonly maxTotalUncompressedBytes: number;
  readonly maxCompressionRatio: number;
}

/** Default safeguards applied when the caller doesn't override them. */
export const DEFAULT_DECOMPRESSION_LIMITS: ResolvedDecompressionLimits = {
  maxEntryUncompressedBytes: 512 * 1024 * 1024,
  maxTotalUncompressedBytes: 1024 * 1024 * 1024,
  maxCompressionRatio: 1000,
};

/** Below this compressed size, the ratio check is skipped — see {@link DecompressionLimits}. */
const RATIO_CHECK_MIN_COMPRESSED_BYTES = 64;

/**
 * `DecompressionLimits` plus the sentinel `false` to disable the guard.
 * Anything else is treated as "use defaults".
 */
export type DecompressionLimitsInput = DecompressionLimits | false | undefined;

/** Returns null when the guard is disabled. */
export function resolveDecompressionLimits(
  input: DecompressionLimitsInput,
): ResolvedDecompressionLimits | null {
  if (input === false) return null;
  if (!input) return DEFAULT_DECOMPRESSION_LIMITS;
  return {
    maxEntryUncompressedBytes:
      input.maxEntryUncompressedBytes ?? DEFAULT_DECOMPRESSION_LIMITS.maxEntryUncompressedBytes,
    maxTotalUncompressedBytes:
      input.maxTotalUncompressedBytes ?? DEFAULT_DECOMPRESSION_LIMITS.maxTotalUncompressedBytes,
    maxCompressionRatio:
      input.maxCompressionRatio ?? DEFAULT_DECOMPRESSION_LIMITS.maxCompressionRatio,
  };
}

/**
 * Per-archive byte accounting shared across every read of an archive opened
 * with the given limits. Tracks total inflated bytes so the cap is enforced
 * even when individual entries stay below the per-entry ceiling.
 */
export interface DecompressionBudget {
  readonly limits: ResolvedDecompressionLimits;
  totalInflated: number;
}

export function createBudget(limits: ResolvedDecompressionLimits): DecompressionBudget {
  return { limits, totalInflated: 0 };
}

/**
 * Per-entry inflate cap, accounting for both the absolute per-entry bound and
 * the ratio-based bound. Returns the smaller of the two — whichever fires
 * first stops the inflate loop.
 */
export function entryInflateCap(budget: DecompressionBudget, compressedSize: number): number {
  const { maxEntryUncompressedBytes, maxCompressionRatio } = budget.limits;
  if (compressedSize < RATIO_CHECK_MIN_COMPRESSED_BYTES) {
    return maxEntryUncompressedBytes;
  }
  const ratioCap = compressedSize * maxCompressionRatio;
  return Math.min(maxEntryUncompressedBytes, ratioCap);
}

/** Throw if the declared central-directory totals already exceed the limits. */
export function checkDeclaredTotals(
  budget: DecompressionBudget,
  declaredEntries: ReadonlyArray<{ path: string; compSize: number; uncompSize: number }>,
): void {
  let declaredTotal = 0;
  for (const entry of declaredEntries) {
    declaredTotal += entry.uncompSize;
    if (entry.uncompSize > budget.limits.maxEntryUncompressedBytes) {
      throw new OpenXmlDecompressionBombError(
        `openZip: entry "${entry.path}" declares ${entry.uncompSize} uncompressed bytes,` +
          ` exceeding the ${budget.limits.maxEntryUncompressedBytes}-byte per-entry limit` +
          ` (decompression-bomb guard).`,
      );
    }
    if (
      entry.compSize >= RATIO_CHECK_MIN_COMPRESSED_BYTES &&
      entry.uncompSize > entry.compSize * budget.limits.maxCompressionRatio
    ) {
      throw new OpenXmlDecompressionBombError(
        `openZip: entry "${entry.path}" declares ratio ${(entry.uncompSize / entry.compSize).toFixed(1)}x` +
          ` (${entry.uncompSize}/${entry.compSize}), exceeding the ${budget.limits.maxCompressionRatio}x` +
          ` per-entry limit (decompression-bomb guard).`,
      );
    }
  }
  if (declaredTotal > budget.limits.maxTotalUncompressedBytes) {
    throw new OpenXmlDecompressionBombError(
      `openZip: declared total uncompressed size ${declaredTotal} bytes exceeds the` +
        ` ${budget.limits.maxTotalUncompressedBytes}-byte archive limit (decompression-bomb guard).`,
    );
  }
}

/**
 * Record `bytes` against the global budget. Throws when the running total
 * crosses {@link ResolvedDecompressionLimits.maxTotalUncompressedBytes}. Called
 * from both sync and streaming inflate code paths.
 */
export function recordInflated(budget: DecompressionBudget, path: string, bytes: number): void {
  budget.totalInflated += bytes;
  if (budget.totalInflated > budget.limits.maxTotalUncompressedBytes) {
    throw new OpenXmlDecompressionBombError(
      `openZip: archive-wide inflated size exceeded ${budget.limits.maxTotalUncompressedBytes} bytes` +
        ` while reading "${path}" (decompression-bomb guard).`,
    );
  }
}

/** Build the message thrown when a single entry exceeds its cap mid-inflate. */
export function entryOverflowError(path: string, cap: number): OpenXmlDecompressionBombError {
  return new OpenXmlDecompressionBombError(
    `openZip: inflated size of "${path}" exceeded ${cap} bytes (decompression-bomb guard).`,
  );
}
