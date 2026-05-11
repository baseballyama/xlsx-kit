// ZIP64 entry-count overflow handling. fflate's `Zip` writer always
// emits a plain ZIP32 EOCD; our writer post-processes the assembled
// archive to splice in a ZIP64 EOCD record + locator when the entry
// count exceeds 65535 and patches the EOCD entry-count fields with
// the 0xFFFF sentinel. The reader detects the sentinel and falls back
// to fflate's `unzipSync` for ZIP64-aware central-directory parsing.

import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node';
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

    // Reader round-trip — falls back to fflate's unzipSync for ZIP64.
    const archive = await openZip(fromBuffer(bytes));
    expect(archive.list().length).toBe(ENTRIES);
    expect(archive.read('f0.txt')).toEqual(payload);
    expect(archive.read(`f${ENTRIES - 1}.txt`)).toEqual(payload);
    archive.close();
  }, /* timeout */ 180_000);
});
