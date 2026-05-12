// ZIP64 entry-count overflow handling. fflate's `Zip` writer always
// emits a plain ZIP32 EOCD; our writer post-processes the assembled
// archive to splice in a ZIP64 EOCD record + locator when the entry
// count exceeds 65535 and patches the EOCD entry-count fields with
// the 0xFFFF sentinel. The reader parses the ZIP64 EOCD + Zip64
// Extended Information extra field directly so the decompression-bomb
// guards stay in effect on ZIP64 archives.

import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node';
import { OpenXmlIoError } from '../../src/utils/exceptions';
import { openZip } from '../../src/zip/reader';
import { createZipWriter } from '../../src/zip/writer';

describe('ZIP32 / ZIP64 entry-count limit', () => {
  it('round-trips 60_000 entries (well under the 65535 cap)', async () => {
    const ENTRIES = 60_000;
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    const payload = new TextEncoder().encode('x');
    for (let i = 0; i < ENTRIES; i++) {
      await writer.addEntry(`f${i}.txt`, payload, { compress: false });
    }
    await writer.finalize();

    const archive = await openZip(fromBuffer(sink.result()));
    expect(archive.list().length).toBe(ENTRIES);
    expect(archive.read('f0.txt')).toEqual(payload);
    expect(archive.read(`f${ENTRIES - 1}.txt`)).toEqual(payload);
    archive.close();
  }, /* timeout */ 120_000);

  it('round-trips 70_000 entries via spliced-in ZIP64 EOCD record + locator', async () => {
    const ENTRIES = 70_000;
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    const payload = new TextEncoder().encode('x');
    for (let i = 0; i < ENTRIES; i++) {
      await writer.addEntry(`f${i}.txt`, payload, { compress: false });
    }
    await writer.finalize();

    const bytes = sink.result();

    // Sanity-check the patch: EOCD entry counts should be the 0xFFFF
    // sentinel, signalling readers to consult the ZIP64 record.
    // Locate the EOCD signature scanning back from the end.
    let eocdOff = -1;
    for (let p = bytes.length - 22; p >= 0; p--) {
      if (
        bytes[p] === 0x50 &&
        bytes[p + 1] === 0x4b &&
        bytes[p + 2] === 0x05 &&
        bytes[p + 3] === 0x06
      ) {
        eocdOff = p;
        break;
      }
    }
    expect(eocdOff).toBeGreaterThan(0);
    expect((bytes[eocdOff + 8] ?? 0) | ((bytes[eocdOff + 9] ?? 0) << 8)).toBe(0xffff);
    expect((bytes[eocdOff + 10] ?? 0) | ((bytes[eocdOff + 11] ?? 0) << 8)).toBe(0xffff);

    // Reader round-trip — the random-access reader parses ZIP64 natively.
    const archive = await openZip(fromBuffer(bytes));
    expect(archive.list().length).toBe(ENTRIES);
    expect(archive.read('f0.txt')).toEqual(payload);
    expect(archive.read(`f${ENTRIES - 1}.txt`)).toEqual(payload);
    archive.close();
  }, /* timeout */ 180_000);

  it('fails closed when an archive declares ZIP64 but the ZIP64 EOCD record is missing', async () => {
    // Build a valid small archive then patch the EOCD entry-count fields to
    // 0xFFFF sentinels WITHOUT splicing in the matching ZIP64 EOCD record.
    // The reader must refuse the file rather than fall back to unzipSync —
    // a malformed ZIP64 archive shouldn't get the unbounded inflate
    // treatment.
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('a.txt', new TextEncoder().encode('hi'), { compress: false });
    await writer.finalize();
    const bytes = sink.result();

    // Locate the EOCD and corrupt its entry-count fields.
    let eocdOff = -1;
    for (let p = bytes.length - 22; p >= 0; p--) {
      if (
        bytes[p] === 0x50 &&
        bytes[p + 1] === 0x4b &&
        bytes[p + 2] === 0x05 &&
        bytes[p + 3] === 0x06
      ) {
        eocdOff = p;
        break;
      }
    }
    expect(eocdOff).toBeGreaterThan(0);
    const patched = new Uint8Array(bytes);
    patched[eocdOff + 8] = 0xff;
    patched[eocdOff + 9] = 0xff;
    patched[eocdOff + 10] = 0xff;
    patched[eocdOff + 11] = 0xff;

    await expect(openZip(fromBuffer(patched))).rejects.toBeInstanceOf(OpenXmlIoError);
  });
});
