// ZIP write layer. Per docs/plan/03-foundations.md §2.2 / docs/plan/06-
// streaming.md §3.4: streaming-deflate via fflate's `Zip` + per-entry
// `ZipDeflate` / `ZipPassThrough` so the writer never holds the whole
// archive in memory. Each addEntry pushes its bytes through the deflate
// stream and the resulting ZIP chunks land on the sink one at a time —
// the buffered `toBytes()` sink concatenates them on finish, while a
// streaming sink can flush them as they arrive.

import { Zip, ZipDeflate, ZipPassThrough } from 'fflate';
import type { XlsxSink } from '../io/sink';
import { OpenXmlIoError, OpenXmlNotImplementedError } from '../utils/exceptions';

// ZIP32 caps the End-of-Central-Directory entry-count fields at 16 bits.
// fflate's writer doesn't emit a ZIP64 record when this overflows, which
// would produce an archive that readers truncate to (count % 65536) on
// the central directory. We surface that as a clear error so callers
// know to split their archives. xlsx files in the wild are several
// orders of magnitude below this limit.
const ZIP32_MAX_ENTRIES = 0xffff;

export interface ZipWriter {
  /**
   * Stage an entry. Buffered: bytes are held until {@link finalize}.
   * Streams (ReadableStream) will be accepted once the streaming writer
   * lands; passing one today throws.
   *
   * `compress` defaults to `true`. Pass `false` for already-compressed
   * payloads (PNG/JPEG/zip-as-binary content like vbaProject.bin) so
   * we don't pay deflate costs for no gain.
   */
  addEntry(path: string, bytes: Uint8Array | ReadableStream<Uint8Array>, opts?: { compress?: boolean }): Promise<void>;

  /**
   * Build the central directory and flush all bytes through the sink.
   * Idempotent; subsequent calls resolve to the same payload.
   */
  finalize(): Promise<Uint8Array>;
}

/**
 * ZIP writer backed by fflate's streaming `Zip` class. Entries are
 * pushed through `ZipDeflate` / `ZipPassThrough` streams as they arrive,
 * so peak memory stays at the size of the in-flight entry plus the
 * output buffer rather than the full archive.
 */
export function createZipWriter(sink: XlsxSink): ZipWriter {
  if (!sink.toBytes) {
    throw new OpenXmlIoError('createZipWriter: sink does not expose a buffered toBytes() factory');
  }
  const writer = sink.toBytes();
  let finalised: Promise<Uint8Array> | undefined;
  let endCalled = false;
  const seen = new Set<string>();
  const errors: Error[] = [];
  let zipFinishResolve: (() => void) | undefined;
  const zipFinishPromise = new Promise<void>((resolve) => {
    zipFinishResolve = resolve;
  });

  const zip = new Zip((err, chunk, final) => {
    if (err) {
      errors.push(err instanceof Error ? err : new Error(String(err)));
      return;
    }
    // ZipDeflate emits an empty trailer chunk on the final callback even
    // when there are no bytes; guard against pushing an undefined chunk.
    if (chunk && chunk.byteLength > 0) writer.write(chunk);
    if (final && zipFinishResolve) {
      zipFinishResolve();
      zipFinishResolve = undefined;
    }
  });

  return {
    async addEntry(path, bytes, opts) {
      if (finalised !== undefined) {
        throw new OpenXmlIoError('createZipWriter: addEntry after finalize');
      }
      if (!(bytes instanceof Uint8Array)) {
        throw new OpenXmlIoError(
          'createZipWriter: ReadableStream entries are not yet supported (deferred to streaming writer)',
        );
      }
      if (seen.has(path)) {
        throw new OpenXmlIoError(`createZipWriter: duplicate entry "${path}"`);
      }
      if (seen.size >= ZIP32_MAX_ENTRIES) {
        throw new OpenXmlNotImplementedError(
          `createZipWriter: archive would exceed the ZIP32 entry limit (${ZIP32_MAX_ENTRIES}). ZIP64 write is not supported by the underlying deflate library; split the archive into multiple files.`,
        );
      }
      seen.add(path);
      const compress = opts?.compress ?? true;
      const file = compress ? new ZipDeflate(path) : new ZipPassThrough(path);
      try {
        zip.add(file);
        file.push(bytes, /* final */ true);
      } catch (cause) {
        throw new OpenXmlIoError(`createZipWriter: failed to add entry "${path}"`, { cause });
      }
      if (errors.length > 0) {
        throw new OpenXmlIoError('createZipWriter: stream error during addEntry', { cause: errors[0] });
      }
    },

    async finalize() {
      if (finalised !== undefined) return finalised;
      finalised = (async () => {
        try {
          if (!endCalled) {
            zip.end();
            endCalled = true;
          }
        } catch (cause) {
          throw new OpenXmlIoError('createZipWriter: failed to finalize zip archive', { cause });
        }
        await zipFinishPromise;
        if (errors.length > 0) {
          throw new OpenXmlIoError('createZipWriter: stream error during finalize', { cause: errors[0] });
        }
        return writer.finish();
      })();
      return finalised;
    },
  };
}
