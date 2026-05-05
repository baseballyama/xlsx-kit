// ZIP layer public surface. Reader and writer are memory-mode for now;
// streaming variants land in subsequent /loop turns (see
// docs/plan/03-foundations.md §2 and docs/plan/06-streaming.md).

export type { ZipArchive } from './reader';
export { openZip } from './reader';
export type { StreamingEntryWriter, ZipWriter } from './writer';
export { createZipWriter } from './writer';
