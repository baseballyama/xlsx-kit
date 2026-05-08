// Environment-neutral byte I/O surface: byte-level Source / Sink types
// plus browser-safe helpers (Blob / Response / ReadableStream / ArrayBuffer
// adapters). Format-specific load / save (xlsx, future docx, pptx) live
// under their respective subpaths and consume these types.

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
