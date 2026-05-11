// ZIP read layer.
//
// `openZip(source)` walks the central directory once and inflates each entry on
// demand inside `read(path)` (see `./random-access-reader.ts`). That keeps peak
// memory at compressed-archive size + per-entry inflate scratch, instead of
// holding every uncompressed entry resident at once the way the old `unzipSync`
// shortcut did. The fallback path through fflate's `unzipSync` is preserved for
// ZIP64 / non-standard archives.

import type { XlsxSource } from '../io/source';
import { OpenXmlIoError, OpenXmlNotImplementedError } from '../utils/exceptions';
import type { DecompressionLimits } from './decompression-guard';
import { openRandomAccessArchive } from './random-access-reader';

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

const isCfbCompoundDocument = (bytes: Uint8Array): boolean => {
  if (bytes.length < CFB_MAGIC.length) return false;
  for (let i = 0; i < CFB_MAGIC.length; i++) {
    if (bytes[i] !== CFB_MAGIC[i]) return false;
  }
  return true;
};

export interface ZipArchive {
  /** Sorted list of all entry paths in the archive. */
  list(): string[];
  /** Synchronous read; throws OpenXmlIoError when the path is unknown. */
  read(path: string): Uint8Array;
  /** Promise variant for symmetry with the future streaming reader. */
  readAsync(path: string): Promise<Uint8Array>;
  /**
   * Streaming read: returns the entry's inflated bytes as a Web
   * `ReadableStream<Uint8Array>` chunk-by-chunk. Lets callers (the streaming
   * worksheet iterator, in particular) push the inflated payload through a SAX
   * parser without first materialising it in full — peak memory for a sheet
   * walk drops to the inflate window + SAX state instead of the entire
   * uncompressed worksheet body. Throws OpenXmlIoError when the path is
   * unknown.
   */
  readStream(path: string): ReadableStream<Uint8Array>;
  /** Whether the archive holds an entry at the given path. */
  has(path: string): boolean;
  /** Release the in-memory entry table. Subsequent reads throw. */
  close(): void;
}

/** Options for {@link openZip}. */
export interface OpenZipOptions {
  /**
   * Decompression-bomb safeguards applied while inflating archive entries. The
   * default limits admit any legitimate xlsx and reject pathological archives
   * (extreme compression ratios, gigabyte-scale entries). Pass `false` to
   * disable the guard entirely — only safe when the source is fully trusted.
   * See {@link DecompressionLimits} for the individual knobs.
   */
  decompressionLimits?: DecompressionLimits | false;
}

/**
 * Open a zip archive from any {@link XlsxSource}. Memory-bounded: the source is
 * fully materialised, then handed to fflate.unzipSync to produce a path → bytes
 * map.
 */
export async function openZip(source: XlsxSource, opts: OpenZipOptions = {}): Promise<ZipArchive> {
  let bytes: Uint8Array;
  try {
    bytes = await source.toBytes();
  } catch (cause) {
    throw new OpenXmlIoError('openZip: failed to read source bytes', { cause });
  }

  // Encrypted xlsx files (Excel 2007+ password protection) wrap the real
  // package inside an OLE Compound File Binary container with the magic
  // signature `D0 CF 11 E0 A1 B1 1A E1`. Detect that early and surface a clear
  // "decrypt first" error rather than letting fflate fail with a generic
  // invalid-zip message.
  if (isCfbCompoundDocument(bytes)) {
    throw new OpenXmlNotImplementedError(
      'Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool first.',
    );
  }

  return openRandomAccessArchive(bytes, opts.decompressionLimits);
}
