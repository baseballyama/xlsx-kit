// Node filesystem + Readable / Writable I/O helpers. Per docs/plan/03-
// foundations.md §1.1.
//
// Kept separate from `./node.ts` so the buffer-only entry stays free of
// `node:fs` / `node:stream` imports — important for the
// `ooxml-js/streaming` browser-targeted bundle, which can re-export
// `fromBuffer` / `toBuffer` without dragging Node-only modules into the
// browser surface. Users who want filesystem I/O reach this module
// directly (or through `ooxml-js/node` once that subpath lands).

import { createReadStream, createWriteStream, readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { Readable, Writable } from 'node:stream';
import { OpenXmlIoError } from '../utils/exceptions';
import type { BufferedSinkWriter, XlsxSink } from './sink';
import type { XlsxSource } from './source';

/**
 * Wrap a filesystem path as an XlsxSource. `toBytes` reads the whole
 * file into memory; `toStream` opens a `fs.createReadStream` and bridges
 * it to a Web {@link ReadableStream} via `Readable.toWeb` so the ZIP
 * reader can iterate without loading the entire xlsx up front.
 */
export function fromFile(path: string): XlsxSource {
  if (typeof path !== 'string' || path.length === 0) {
    throw new OpenXmlIoError('fromFile expects a non-empty path string');
  }
  return {
    async toBytes() {
      try {
        return new Uint8Array(await readFile(path));
      } catch (cause) {
        throw new OpenXmlIoError(`fromFile: failed to read "${path}"`, { cause });
      }
    },
    toStream() {
      const nodeStream = createReadStream(path);
      return Readable.toWeb(nodeStream) as unknown as ReadableStream<Uint8Array>;
    },
  };
}

/**
 * Synchronous variant of {@link fromFile}. Convenience for tooling /
 * scripts where the cost of `await fs.readFile` outweighs the
 * ergonomic gain. The returned source's `toBytes` resolves immediately
 * with the bytes already in memory.
 */
export function fromFileSync(path: string): XlsxSource {
  if (typeof path !== 'string' || path.length === 0) {
    throw new OpenXmlIoError('fromFileSync expects a non-empty path string');
  }
  let bytes: Uint8Array;
  try {
    bytes = new Uint8Array(readFileSync(path));
  } catch (cause) {
    throw new OpenXmlIoError(`fromFileSync: failed to read "${path}"`, { cause });
  }
  return {
    async toBytes() {
      return bytes;
    },
    toStream() {
      return new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(bytes);
          controller.close();
        },
      });
    },
  };
}

/**
 * Filesystem sink. Each `write(chunk)` call streams the bytes to disk
 * via `fs.createWriteStream`, so the streaming ZIP backend can flush
 * chunks as they arrive instead of buffering the whole archive.
 * `result()` returns the destination path; `finish()` resolves with
 * the on-disk bytes for callers that want to inspect them after the
 * write (mirrors the toBuffer / toBlob shape).
 */
export function toFile(path: string): XlsxSink & { toBytes(): BufferedSinkWriter; result(): string } {
  if (typeof path !== 'string' || path.length === 0) {
    throw new OpenXmlIoError('toFile expects a non-empty path string');
  }
  let stream: ReturnType<typeof createWriteStream> | undefined;
  let finalised: Promise<Uint8Array> | undefined;
  let pendingError: Error | undefined;

  const ensureStream = (): NonNullable<typeof stream> => {
    if (!stream) {
      stream = createWriteStream(path);
      stream.on('error', (err) => {
        pendingError = err instanceof Error ? err : new Error(String(err));
      });
    }
    return stream;
  };

  return {
    toBytes(): BufferedSinkWriter {
      return {
        write(chunk: Uint8Array): void {
          if (finalised !== undefined) throw new OpenXmlIoError(`toFile sink: write after finish ("${path}")`);
          if (!(chunk instanceof Uint8Array)) {
            throw new OpenXmlIoError(`toFile sink: chunk is not a Uint8Array ("${path}")`);
          }
          if (pendingError) throw new OpenXmlIoError(`toFile sink: write error on "${path}"`, { cause: pendingError });
          ensureStream().write(chunk);
        },
        async finish(): Promise<Uint8Array> {
          if (finalised) return finalised;
          finalised = (async () => {
            const s = ensureStream();
            await new Promise<void>((resolve, reject) => {
              s.end((err?: Error | null) => (err ? reject(err) : resolve()));
            });
            if (pendingError) throw new OpenXmlIoError(`toFile sink: write error on "${path}"`, { cause: pendingError });
            try {
              return new Uint8Array(await readFile(path));
            } catch (cause) {
              throw new OpenXmlIoError(`toFile sink: failed to re-read "${path}"`, { cause });
            }
          })();
          return finalised;
        },
      };
    },
    result(): string {
      return path;
    },
  };
}

/**
 * Wrap a Node.js {@link Readable} as an XlsxSource. `toBytes` consumes
 * the entire stream synchronously (collecting chunks); `toStream`
 * bridges to a Web ReadableStream via `Readable.toWeb` so the ZIP
 * reader can pull chunks lazily.
 */
export function fromReadable(readable: Readable): XlsxSource {
  if (!(readable instanceof Readable)) {
    throw new OpenXmlIoError('fromReadable expects a Node Readable');
  }
  let bytes: Promise<Uint8Array> | undefined;
  return {
    async toBytes() {
      if (bytes) return bytes;
      bytes = (async () => {
        const chunks: Uint8Array[] = [];
        for await (const c of readable) {
          chunks.push(c instanceof Uint8Array ? c : new Uint8Array(c));
        }
        let total = 0;
        for (const c of chunks) total += c.byteLength;
        const out = new Uint8Array(total);
        let off = 0;
        for (const c of chunks) {
          out.set(c, off);
          off += c.byteLength;
        }
        return out;
      })();
      return bytes;
    },
    toStream() {
      return Readable.toWeb(readable) as unknown as ReadableStream<Uint8Array>;
    },
  };
}

/**
 * Wrap a Node.js {@link Writable} as an XlsxSink. Each write streams
 * directly to the underlying writable; `result()` returns the writable
 * itself for downstream chaining.
 */
export function toWritable(writable: Writable): XlsxSink & { toBytes(): BufferedSinkWriter; result(): Writable } {
  if (!(writable instanceof Writable)) {
    throw new OpenXmlIoError('toWritable expects a Node Writable');
  }
  let finalised: Promise<Uint8Array> | undefined;
  let pendingError: Error | undefined;
  writable.on('error', (err) => {
    pendingError = err instanceof Error ? err : new Error(String(err));
  });

  return {
    toBytes(): BufferedSinkWriter {
      return {
        write(chunk: Uint8Array): void {
          if (finalised !== undefined) throw new OpenXmlIoError('toWritable sink: write after finish');
          if (!(chunk instanceof Uint8Array)) throw new OpenXmlIoError('toWritable sink: chunk is not a Uint8Array');
          if (pendingError) throw new OpenXmlIoError('toWritable sink: write error', { cause: pendingError });
          writable.write(chunk);
        },
        async finish(): Promise<Uint8Array> {
          if (finalised) return finalised;
          finalised = (async () => {
            await new Promise<void>((resolve, reject) => {
              writable.end((err?: Error | null) => (err ? reject(err) : resolve()));
            });
            if (pendingError) throw new OpenXmlIoError('toWritable sink: write error', { cause: pendingError });
            return new Uint8Array(0);
          })();
          return finalised;
        },
      };
    },
    result(): Writable {
      return writable;
    },
  };
}
