// Node-only public entry. Filesystem / Readable / Writable bridges plus
// Buffer-based source/sink. Format-specific load/save (xlsx) lives under
// its own subpath (`ooxml-js/xlsx/io`).

export { fromBuffer, toBuffer } from './io/node';
export { fromFile, fromFileSync, fromReadable, toFile, toWritable } from './io/node-fs';
