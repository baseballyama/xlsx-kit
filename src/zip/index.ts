// ZIP layer public surface. Reader and writer are memory-mode for now;
// streaming variants live alongside.

export type {
  DecompressionLimits,
  ResolvedDecompressionLimits,
} from './decompression-guard';
export { DEFAULT_DECOMPRESSION_LIMITS } from './decompression-guard';
export type { OpenZipOptions, ZipArchive } from './reader';
export { openZip } from './reader';
export type { StreamingEntryWriter, ZipWriter } from './writer';
export { createZipWriter } from './writer';
