// Public streaming entry point. Per docs/plan/01-architecture.md §3 +
// docs/plan/06-streaming.md: a narrow subpath (`openxml-js/streaming`)
// so callers that only need read-only iter / write-only append don't
// pull in the full workbook model. The size-limit gate (≤80KB min+gz)
// guards against accidentally re-importing core modules from here.

export {
  loadWorkbookStream,
  type IterRowsOptions,
  type ReadOnlyCell,
  type ReadOnlyWorkbook,
  type ReadOnlyWorksheet,
} from './read-only';

export {
  createWriteOnlyWorkbook,
  type WriteOnlyOptions,
  type WriteOnlyRowItem,
  type WriteOnlyStyle,
  type WriteOnlyWorkbook,
  type WriteOnlyWorksheet,
} from './write-only';

// Environment-neutral I/O — `fromBuffer` works under Node and any
// runtime that polyfills Buffer; the browser-specific helpers are
// always-callable from Node 18+ since Blob / File / fetch are global.
// Filesystem-bound helpers (fromFile / toFile / fromReadable / toWritable)
// live behind the `openxml-js/node` subpath to keep this entry browser-
// safe — importing `node:fs` here would fail under Vite / webpack.
export { fromBuffer, toBuffer } from '../io/node';
export {
  fromArrayBuffer,
  fromBlob,
  fromFile,
  fromResponse,
  fromStream,
  toArrayBuffer,
  toBlob,
} from '../io/browser';
export type { BufferedSinkWriter, XlsxSink } from '../io/sink';
export type { XlsxSource } from '../io/source';

// Error surface — callers need to type-narrow on these.
export {
  OpenXmlError,
  OpenXmlIoError,
  OpenXmlSchemaError,
  OpenXmlInvalidWorkbookError,
  OpenXmlNotImplementedError,
} from '../utils/exceptions';
