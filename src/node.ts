// Node-only public entry. Per docs/plan/03-foundations.md §1.1 and
// docs/plan/11-build-publish.md §1.3 — `openxml-js/node` keeps the
// filesystem / Readable / Writable surface separate from the
// browser-safe core so users importing it get the right type
// definitions and bundlers don't trip on `node:fs` in browser
// targets.

export { fromBuffer, toBuffer } from './io/node';
export { fromFile, fromFileSync, toFile, fromReadable, toWritable } from './io/node-fs';
export type { BufferedSinkWriter, XlsxSink } from './io/sink';
export type { XlsxSource } from './io/source';

// Convenience re-exports of the high-level entry points so a Node user
// can `import { loadWorkbook, fromFile } from 'openxml-js/node'` without
// pulling from two paths.
export { loadWorkbook, type LoadOptions } from './public/load';
export { saveWorkbook, workbookToBytes, type SaveOptions } from './public/save';
