// ZIP layer public surface. Reader is memory-mode-only for now; streaming
// reader and full writer land in subsequent /loop turns (see
// docs/plan/03-foundations.md §2 and docs/plan/06-streaming.md).

export type { ZipArchive } from './reader';
export { openZip } from './reader';
