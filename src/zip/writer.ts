// ZIP write layer. Phase-1 §2: in-memory deflate via fflate.zipSync.
//
// The plan calls for the eventual writer to use fflate's streaming `Zip`
// API with per-entry ZipDeflate / ZipPassThrough choice
// (docs/plan/03-foundations.md §2.2). That lands when the write-only
// worksheet path needs it (phase 4). For now a buffered all-at-once
// writer is enough to round-trip xlsx files in tests.

import { type DeflateOptions, type ZipAttributes, zipSync } from 'fflate';
import type { XlsxSink } from '../io/sink';
import { OpenXmlIoError } from '../utils/exceptions';

// fflate accepts per-entry options as the union of ZipAttributes and
// DeflateOptions; `level: 0` selects STORE, otherwise deflate is used.
type EntryOptions = ZipAttributes & DeflateOptions;

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
 * Buffered ZIP writer over an XlsxSink. The sink must expose a buffered
 * `toBytes()` factory; streaming sinks will be supported alongside the
 * streaming writer.
 */
export function createZipWriter(sink: XlsxSink): ZipWriter {
  if (!sink.toBytes) {
    throw new OpenXmlIoError('createZipWriter: sink does not expose a buffered toBytes() factory');
  }
  const writer = sink.toBytes();
  const entries: Record<string, Uint8Array | [Uint8Array, EntryOptions]> = {};
  let finalised: Promise<Uint8Array> | undefined;

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
      if (Object.hasOwn(entries, path)) {
        throw new OpenXmlIoError(`createZipWriter: duplicate entry "${path}"`);
      }
      const compress = opts?.compress ?? true;
      // fflate compression: level 0 = STORE (no compression).
      // Default (omitting `level`) is deflate with reasonable settings.
      entries[path] = compress ? bytes : [bytes, { level: 0 }];
    },

    async finalize() {
      if (finalised !== undefined) return finalised;
      finalised = (async () => {
        let zipped: Uint8Array;
        try {
          zipped = zipSync(entries);
        } catch (cause) {
          throw new OpenXmlIoError('createZipWriter: failed to assemble zip archive', { cause });
        }
        writer.write(zipped);
        return writer.finish();
      })();
      return finalised;
    },
  };
}
