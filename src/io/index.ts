// I/O surface: byte-level Source / Sink types, browser-safe byte helpers
// (Blob / Response / ReadableStream / ArrayBuffer adapters), and the
// xlsx load / save / serialise entry points.

export type { BufferedSinkWriter, XlsxSink } from './sink';
export type { XlsxSource } from './source';
export {
  fromArrayBuffer,
  fromBlob,
  fromResponse,
  fromStream,
  toArrayBuffer,
  toBlob,
} from './browser';
export type { LoadOptions } from './load';
export { loadWorkbook } from './load';
export type { SaveOptions } from './save';
export { saveWorkbook, workbookToBytes } from './save';
