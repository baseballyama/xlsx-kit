// Browser-target I/O helpers. Phase-1 §1: synchronous in-memory paths.
// Response / streaming-WritableStream bridges land alongside the
// streaming ZIP work in §2.
//
// Node 18+ provides Blob / File / FormData / fetch / Web Streams natively,
// so this module is also exercised by the Node-hosted vitest runner.

import { OpenXmlIoError } from '../utils/exceptions';
import type { BufferedSinkWriter, XlsxSink } from './sink';
import type { XlsxSource } from './source';

/**
 * Wrap a Blob (or File, since File extends Blob) as an XlsxSource. The
 * source materialises bytes lazily on the first {@link XlsxSource.toBytes}
 * call.
 */
export function fromBlob(blob: Blob): XlsxSource {
  if (!(blob instanceof Blob)) {
    throw new OpenXmlIoError('fromBlob expects a Blob (or File)');
  }
  return {
    async toBytes() {
      try {
        const ab = await blob.arrayBuffer();
        return new Uint8Array(ab);
      } catch (cause) {
        throw new OpenXmlIoError('fromBlob: failed to read blob contents', { cause });
      }
    },
    toStream() {
      return blob.stream();
    },
  };
}

/** Alias kept for symmetry with the plan; File extends Blob so the implementation is shared. */
export const fromFile = fromBlob;

/**
 * Wrap an ArrayBuffer or Uint8Array. The bytes are referenced (Uint8Array
 * input) or wrapped without copy (ArrayBuffer input).
 */
export function fromArrayBuffer(buf: ArrayBuffer | Uint8Array): XlsxSource {
  let bytes: Uint8Array;
  if (buf instanceof Uint8Array) {
    bytes = buf;
  } else if (buf instanceof ArrayBuffer) {
    bytes = new Uint8Array(buf);
  } else {
    throw new OpenXmlIoError('fromArrayBuffer expects an ArrayBuffer or Uint8Array');
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

const collectChunks = (chunks: Uint8Array[]): Uint8Array => {
  let total = 0;
  for (const c of chunks) total += c.byteLength;
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.byteLength;
  }
  return out;
};

/**
 * In-memory Blob sink. Convenience `result()` returns the accumulated
 * payload as a Blob with the supplied MIME type.
 */
export function toBlob(
  mime: string = 'application/octet-stream',
): XlsxSink & { toBytes(): BufferedSinkWriter; result(): Blob } {
  const chunks: Uint8Array[] = [];
  let finalised: Uint8Array | undefined;

  const finalise = (): Uint8Array => {
    if (finalised !== undefined) return finalised;
    finalised = collectChunks(chunks);
    chunks.length = 0;
    return finalised;
  };

  return {
    toBytes(): BufferedSinkWriter {
      return {
        write(chunk: Uint8Array): void {
          if (finalised !== undefined) throw new OpenXmlIoError('toBlob sink: write after finish');
          if (!(chunk instanceof Uint8Array)) throw new OpenXmlIoError('toBlob sink: chunk is not a Uint8Array');
          chunks.push(chunk);
        },
        async finish(): Promise<Uint8Array> {
          return finalise();
        },
      };
    },
    result(): Blob {
      // Copy the bytes' typed-array view into a fresh Blob to detach from
      // the writer's internal buffer.
      const bytes = finalise();
      return new Blob([bytes.slice()], { type: mime });
    },
  };
}

/** In-memory ArrayBuffer sink. */
export function toArrayBuffer(): XlsxSink & { toBytes(): BufferedSinkWriter; result(): ArrayBuffer } {
  const chunks: Uint8Array[] = [];
  let finalised: Uint8Array | undefined;

  const finalise = (): Uint8Array => {
    if (finalised !== undefined) return finalised;
    finalised = collectChunks(chunks);
    chunks.length = 0;
    return finalised;
  };

  return {
    toBytes(): BufferedSinkWriter {
      return {
        write(chunk: Uint8Array): void {
          if (finalised !== undefined) throw new OpenXmlIoError('toArrayBuffer sink: write after finish');
          if (!(chunk instanceof Uint8Array)) throw new OpenXmlIoError('toArrayBuffer sink: chunk is not a Uint8Array');
          chunks.push(chunk);
        },
        async finish(): Promise<Uint8Array> {
          return finalise();
        },
      };
    },
    result(): ArrayBuffer {
      const bytes = finalise();
      // bytes' underlying buffer may be larger than its byteLength (slice
      // semantics on Uint8Array). Materialise an exactly-sized copy.
      const out = new ArrayBuffer(bytes.byteLength);
      new Uint8Array(out).set(bytes);
      return out;
    },
  };
}
