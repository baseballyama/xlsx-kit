import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../../src/io/node';
import { OpenXmlIoError } from '../../../src/utils/exceptions';
import { openZip } from '../../../src/zip/reader';
import { createZipWriter } from '../../../src/zip/writer';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../../reference/openpyxl/openpyxl/tests/data/genuine');
const EMPTY_XLSX = resolve(FIXTURES, 'empty.xlsx');

const utf8 = (s: string): Uint8Array => new TextEncoder().encode(s);

describe('createZipWriter (basic)', () => {
  it('produces a zip readable by openZip with the same bytes', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('hello.txt', utf8('hello'));
    await writer.addEntry('a/b.txt', utf8('nested'));
    await writer.finalize();

    const zip = await openZip(fromBuffer(sink.result()));
    expect(zip.list()).toEqual(['a/b.txt', 'hello.txt']);
    expect(new TextDecoder().decode(zip.read('hello.txt'))).toBe('hello');
    expect(new TextDecoder().decode(zip.read('a/b.txt'))).toBe('nested');
  });

  it('honours compress: false (STORE) — bytes still round-trip', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    const payload = new Uint8Array(1024).fill(0xab);
    await writer.addEntry('xl/media/image1.png', payload, { compress: false });
    await writer.finalize();

    const zip = await openZip(fromBuffer(sink.result()));
    expect(zip.read('xl/media/image1.png')).toEqual(payload);
  });

  it('handles 0-byte entries', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('empty.bin', new Uint8Array(0));
    await writer.finalize();

    const zip = await openZip(fromBuffer(sink.result()));
    expect(zip.read('empty.bin').byteLength).toBe(0);
  });

  it('finalize() is idempotent and yields the same bytes', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('a', utf8('A'));
    const first = await writer.finalize();
    const second = await writer.finalize();
    expect(second).toBe(first);
  });

  it('rejects addEntry after finalize', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('a', utf8('A'));
    await writer.finalize();
    await expect(writer.addEntry('b', utf8('B'))).rejects.toBeInstanceOf(OpenXmlIoError);
  });

  it('rejects duplicate entry paths', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('a', utf8('A'));
    await expect(writer.addEntry('a', utf8('A2'))).rejects.toBeInstanceOf(OpenXmlIoError);
  });

  it('rejects ReadableStream payload (deferred to streaming writer)', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(utf8('hi'));
        controller.close();
      },
    });
    await expect(writer.addEntry('a', stream)).rejects.toBeInstanceOf(OpenXmlIoError);
  });

  it('rejects sinks without a buffered toBytes()', () => {
    const stubSink = {} as Parameters<typeof createZipWriter>[0];
    expect(() => createZipWriter(stubSink)).toThrowError(OpenXmlIoError);
  });
});

describe('createZipWriter (round-trip via openZip)', () => {
  it('every entry of empty.xlsx round-trips byte-identically through the writer', async () => {
    const original = await openZip(fromBuffer(readFileSync(EMPTY_XLSX)));
    const paths = original.list();

    const sink = toBuffer();
    const writer = createZipWriter(sink);
    for (const path of paths) {
      // OOXML xml parts are deflate by default; use STORE for media + vba.
      const compress = !(path.startsWith('xl/media/') || path === 'xl/vbaProject.bin');
      await writer.addEntry(path, original.read(path), { compress });
    }
    await writer.finalize();

    const round = await openZip(fromBuffer(sink.result()));
    expect(round.list()).toEqual(paths);
    for (const path of paths) {
      expect(round.read(path)).toEqual(original.read(path));
    }
  });
});
