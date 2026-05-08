// ZIP write layer. Per docs/plan/03-foundations.md §2.2 / docs/plan/06-
// streaming.md §3.4: streaming-deflate via fflate's `Zip` + per-entry
// `ZipDeflate` / `ZipPassThrough` so the writer never holds the whole
// archive in memory. Each addEntry pushes its bytes through the deflate
// stream and the resulting ZIP chunks land on the sink one at a time —
// the buffered `toBytes()` sink concatenates them on finish, while a
// streaming sink can flush them as they arrive.
//
// ZIP64 (entry count > 65535): fflate's `Zip` emits a plain ZIP32 EOCD
// in all cases, so on finalize we splice in a ZIP64 EOCD record + locator
// when needed via `applyZip64EntryCountPatch`. That keeps the per-entry
// LFH/CDH layout fflate produces (correct as long as no individual size
// or offset overflows 32 bits) and only rewrites the trailing records
// — enough for xlsx, which never approaches single-entry 4 GiB limits.

import { Zip, ZipDeflate, ZipPassThrough } from 'fflate';
import type { XlsxSink } from '../io/sink';
import { OpenXmlIoError } from '../utils/exceptions';
import { applyZip64EntryCountPatch } from './zip64-patch';

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
   * Open a streaming entry. Returns a writer the caller can `write()`
   * chunks to and `end()` to seal the entry. Each chunk pushes through
   * the same fflate `ZipDeflate` / `ZipPassThrough` machinery as
   * `addEntry`, so peak memory stays at one chunk + deflate scratch
   * even for multi-GB worksheets.
   *
   * Sequencing: only one streaming entry may be open at a time —
   * `addEntry` and a second `addStreamingEntry` both throw until the
   * current entry's `end()` resolves.
   */
  addStreamingEntry(path: string, opts?: { compress?: boolean }): StreamingEntryWriter;

  /**
   * Build the central directory and flush all bytes through the sink.
   * Idempotent; subsequent calls resolve to the same payload.
   */
  finalize(): Promise<Uint8Array>;
}

/** Writer handle for a single streaming entry. */
export interface StreamingEntryWriter {
  /** Push a chunk of bytes (already-encoded). Throws after `end()`. */
  write(chunk: Uint8Array): void;
  /** Seal the entry. Subsequent `write()` throws. Idempotent. */
  end(): Promise<void>;
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
  // fflate emits the [CD | EOCD] block in a single ondata call with
  // `final=true`. We capture only that chunk so we can apply the ZIP64
  // patch on finalize; all preceding entry-data chunks stream straight
  // to the sink to preserve the writer's incremental flushing contract.
  let finalChunk: Uint8Array | undefined;
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
    if (chunk && chunk.byteLength > 0) {
      if (final) {
        // Buffer the trailing CD + EOCD block; written after possible patch.
        finalChunk = chunk;
      } else {
        writer.write(chunk);
      }
    }
    if (final && zipFinishResolve) {
      zipFinishResolve();
      zipFinishResolve = undefined;
    }
  });

  let streamingOpen = false;

  const guardAdd = (path: string): void => {
    if (finalised !== undefined) {
      throw new OpenXmlIoError('createZipWriter: addEntry after finalize');
    }
    if (streamingOpen) {
      throw new OpenXmlIoError('createZipWriter: a streaming entry is still open — call end() first');
    }
    if (seen.has(path)) {
      throw new OpenXmlIoError(`createZipWriter: duplicate entry "${path}"`);
    }
  };

  return {
    async addEntry(path, bytes, opts) {
      if (!(bytes instanceof Uint8Array)) {
        throw new OpenXmlIoError(
          'createZipWriter: ReadableStream entries are not yet supported (deferred to streaming writer)',
        );
      }
      guardAdd(path);
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

    addStreamingEntry(path, opts) {
      guardAdd(path);
      seen.add(path);
      streamingOpen = true;
      const compress = opts?.compress ?? true;
      const file = compress ? new ZipDeflate(path) : new ZipPassThrough(path);
      try {
        zip.add(file);
      } catch (cause) {
        streamingOpen = false;
        throw new OpenXmlIoError(`createZipWriter: failed to open streaming entry "${path}"`, { cause });
      }
      let ended = false;
      return {
        write(chunk: Uint8Array): void {
          if (ended) throw new OpenXmlIoError(`createZipWriter: write after end on "${path}"`);
          if (!(chunk instanceof Uint8Array)) {
            throw new OpenXmlIoError(`createZipWriter: streaming entry "${path}" chunk is not a Uint8Array`);
          }
          if (chunk.byteLength === 0) return;
          try {
            file.push(chunk, /* final */ false);
          } catch (cause) {
            throw new OpenXmlIoError(`createZipWriter: failed to push chunk on "${path}"`, { cause });
          }
          if (errors.length > 0) {
            throw new OpenXmlIoError('createZipWriter: stream error during write', { cause: errors[0] });
          }
        },
        async end(): Promise<void> {
          if (ended) return;
          ended = true;
          try {
            file.push(new Uint8Array(0), /* final */ true);
          } catch (cause) {
            throw new OpenXmlIoError(`createZipWriter: failed to end streaming entry "${path}"`, { cause });
          }
          streamingOpen = false;
          if (errors.length > 0) {
            throw new OpenXmlIoError('createZipWriter: stream error during end', { cause: errors[0] });
          }
        },
      };
    },

    async finalize() {
      if (finalised !== undefined) return finalised;
      if (streamingOpen) {
        throw new OpenXmlIoError('createZipWriter: cannot finalize while a streaming entry is open');
      }
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

        // Apply the ZIP64 patch to fflate's [CD | EOCD] tail when the
        // entry count exceeds ZIP32's 16-bit cap, then flush the
        // (possibly patched) tail to the sink.
        if (finalChunk) {
          const patched =
            seen.size > ZIP32_MAX_ENTRIES
              ? applyZip64EntryCountPatch(finalChunk, seen.size)
              : finalChunk;
          writer.write(patched);
        }
        return writer.finish();
      })();
      return finalised;
    },
  };
}
