// ZIP read layer. Phase-1 §2: in-memory deflate via fflate.unzipSync.
//
// `openZip(source)` materialises every entry's bytes up-front. That keeps
// the API simple and is fine for the typical xlsx (≤ tens of MB). The
// streaming path documented in docs/plan/03-foundations.md §2.1 lands when
// the read-only / write-only worksheet modes do (phase 4).

import { unzipSync } from 'fflate';
import type { XlsxSource } from '../io/source';
import { OpenXmlIoError, OpenXmlNotImplementedError } from '../utils/exceptions';

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
  /** Whether the archive holds an entry at the given path. */
  has(path: string): boolean;
  /** Release the in-memory entry table. Subsequent reads throw. */
  close(): void;
}

/**
 * Open a zip archive from any {@link XlsxSource}. Memory-bounded:
 * the source is fully materialised, then handed to fflate.unzipSync to
 * produce a path → bytes map.
 */
export async function openZip(source: XlsxSource): Promise<ZipArchive> {
  let bytes: Uint8Array;
  try {
    bytes = await source.toBytes();
  } catch (cause) {
    throw new OpenXmlIoError('openZip: failed to read source bytes', { cause });
  }

  // Encrypted xlsx files (Excel 2007+ password protection) wrap the real
  // package inside an OLE Compound File Binary container with the magic
  // signature `D0 CF 11 E0 A1 B1 1A E1`. Detect that early and surface a
  // clear "decrypt first" error rather than letting fflate fail with a
  // generic invalid-zip message.
  if (isCfbCompoundDocument(bytes)) {
    throw new OpenXmlNotImplementedError(
      'Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool first.',
    );
  }

  let entries: Record<string, Uint8Array> | undefined;
  try {
    entries = unzipSync(bytes);
  } catch (cause) {
    throw new OpenXmlIoError('openZip: archive is not a valid zip', { cause });
  }

  let live = true;
  const ensureLive = (): Record<string, Uint8Array> => {
    if (!live || entries === undefined) {
      throw new OpenXmlIoError('openZip: archive is closed');
    }
    return entries;
  };

  const readEntry = (path: string): Uint8Array => {
    const e = ensureLive();
    const found = e[path];
    if (found === undefined) {
      throw new OpenXmlIoError(`openZip: no entry at "${path}"`);
    }
    return found;
  };

  return {
    list(): string[] {
      const e = ensureLive();
      return Object.keys(e).sort();
    },
    has(path: string): boolean {
      if (!live || entries === undefined) return false;
      return Object.hasOwn(entries, path);
    },
    read(path: string): Uint8Array {
      return readEntry(path);
    },
    async readAsync(path: string): Promise<Uint8Array> {
      return readEntry(path);
    },
    close(): void {
      live = false;
      entries = undefined;
    },
  };
}
