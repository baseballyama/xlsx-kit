// ZIP layer public surface. Reader and writer are memory-mode for now;
// streaming variants live alongside.

export type { ZipArchive } from './reader';
export { openZip } from './reader';
export type { StreamingEntryWriter, ZipWriter } from './writer';
export { createZipWriter } from './writer';
