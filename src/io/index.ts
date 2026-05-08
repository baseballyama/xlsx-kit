// Environment-neutral I/O surface: byte-level types (XlsxSource / XlsxSink)
// plus the high-level loadWorkbook / saveWorkbook / workbookToBytes entries.
// Node-only and browser-only helpers live in node.ts / browser.ts and are
// reached via `openxml-js/node` / `openxml-js/io/browser`.

export type { BufferedSinkWriter, XlsxSink } from './sink';
export type { XlsxSource } from './source';
export type { LoadOptions } from '../public/load';
export { loadWorkbook } from '../public/load';
export type { SaveOptions } from '../public/save';
export { saveWorkbook, workbookToBytes } from '../public/save';
