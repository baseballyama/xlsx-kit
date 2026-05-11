// Node filesystem + Readable / Writable I/O helpers.
//
// Kept separate from `./node.ts` so the buffer-only entry stays free of
// `node:fs` / `node:stream` imports — important for the `xlsx-kit/streaming`
// browser-targeted bundle, which can re-export `fromBuffer` / `toBuffer`
// without dragging Node-only modules into the browser surface. Users who want
// filesystem I/O reach this module directly (or through `xlsx-kit/node` once
// that subpath lands).

import { createReadStream, createWriteStream, readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { once } from 'node:events';
import { Readable, Writable } from 'node:stream';
import { OpenXmlIoError } from '../utils/exceptions';
import type { BufferedSinkWriter, XlsxSink } from './sink';
import type { XlsxSource } from './source';

const EMPTY_BYTES = new Uint8Array(0);

/**
 * Wrap a filesystem path as an XlsxSource. `toBytes` reads the whole file into
 * memory; `toStream` opens a `fs.createReadStream` and bridges it to a Web
 * {@link ReadableStream} via `Readable.toWeb` so the ZIP reader can iterate
 * without loading the entire xlsx up front.
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
 * Synchronous variant of {@link fromFile}. Convenience for tooling / scripts
 * where the cost of `await fs.readFile` outweighs the ergonomic gain. The
 * returned source's `toBytes` resolves immediately with the bytes already in
 * memory.
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
 * Filesystem sink. Each `write(chunk)` call streams the bytes to disk via
 * `fs.createWriteStream`, honouring backpressure: when the underlying writable
 * signals it is full, `write` awaits a `drain` event before returning. That
 * keeps peak memory bounded to the writable's `highWaterMark` (default 16 KB)
 * even when the producer races ahead.
 *
 * `result()` returns the destination path; `finish()` resolves with an empty
 * `Uint8Array` once the stream has flushed. Callers that need the on-disk
 * bytes should `readFile()` the returned path themselves — re-reading inside
 * `finish()` would defeat the "streamed to disk, never resident" guarantee.
 */
export function toFile(path: string): XlsxSink & { toBytes(): BufferedSinkWriter; result(): string } {
  if (typeof path !== 'string' || path.length === 0) {
    throw new OpenXmlIoError('toFile expects a non-empty path string');
  }
  let stream: ReturnType<typeof createWriteStream> | undefined;
  let finalised: Promise<Uint8Array> | undefined;
  let pendingError: Error | undefined;
  // Backpressure: chain each write off the previous so a slow disk pulls the
  // producer back. We deliberately don't expose this on the synchronous-style
  // `write()` API — the writer doesn't await — but every write enqueues a
  // promise that `finish()` drains before resolving.
  let writeQueue: Promise<void> = Promise.resolve();

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
          const s = ensureStream();
          const ok = s.write(chunk);
          if (!ok) {
            // Producer is outpacing the disk — chain the next operation behind
            // a `drain` so subsequent writes stage their chunks instead of
            // piling up inside the writable's internal buffer.
            writeQueue = writeQueue.then(() => once(s, 'drain').then(() => undefined));
          }
        },
        async finish(): Promise<Uint8Array> {
          if (finalised) return finalised;
          finalised = (async () => {
            const s = ensureStream();
            await writeQueue;
            await new Promise<void>((resolve, reject) => {
              s.end((err?: Error | null) => (err ? reject(err) : resolve()));
            });
            if (pendingError) throw new OpenXmlIoError(`toFile sink: write error on "${path}"`, { cause: pendingError });
            return EMPTY_BYTES;
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
 * Wrap a Node.js {@link Readable} as an XlsxSource. `toBytes` consumes the
 * entire stream synchronously (collecting chunks); `toStream` bridges to a Web
 * ReadableStream via `Readable.toWeb` so the ZIP reader can pull chunks lazily.
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
 * Wrap a Node.js {@link Writable} as an XlsxSink. Each write streams directly
 * to the underlying writable, honouring backpressure: when the writable's
 * internal buffer is full, the next sink-level write parks behind a `drain`
 * event so memory pressure is bounded by the writable's `highWaterMark`
 * rather than the producer's pace. `result()` returns the writable itself for
 * downstream chaining.
 */
export function toWritable(writable: Writable): XlsxSink & { toBytes(): BufferedSinkWriter; result(): Writable } {
  if (!(writable instanceof Writable)) {
    throw new OpenXmlIoError('toWritable expects a Node Writable');
  }
  let finalised: Promise<Uint8Array> | undefined;
  let pendingError: Error | undefined;
  let writeQueue: Promise<void> = Promise.resolve();
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
          const ok = writable.write(chunk);
          if (!ok) {
            writeQueue = writeQueue.then(() => once(writable, 'drain').then(() => undefined));
          }
        },
        async finish(): Promise<Uint8Array> {
          if (finalised) return finalised;
          finalised = (async () => {
            await writeQueue;
            await new Promise<void>((resolve, reject) => {
              writable.end((err?: Error | null) => (err ? reject(err) : resolve()));
            });
            if (pendingError) throw new OpenXmlIoError('toWritable sink: write error', { cause: pendingError });
            return EMPTY_BYTES;
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
