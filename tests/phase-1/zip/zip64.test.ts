// ZIP64 limit handling — fflate's `Zip` writer doesn't emit a ZIP64
// end-of-central-directory record when the entry count overflows the
// 16-bit ZIP32 cap (65535 entries). Rather than silently produce a
// truncated archive, our writer detects the overflow and throws
// OpenXmlNotImplementedError. xlsx files are several orders of magnitude
// below the cap in practice, so the limit is documented + guarded.

import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../../src/io/node';
import { openZip } from '../../../src/zip/reader';
import { createZipWriter } from '../../../src/zip/writer';
import { OpenXmlNotImplementedError } from '../../../src/utils/exceptions';

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

  it('rejects archives that would exceed the 65535-entry ZIP32 cap', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    const payload = new TextEncoder().encode('x');
    // Add up to the cap.
    for (let i = 0; i < 0xffff; i++) {
      await writer.addEntry(`f${i}.txt`, payload, { compress: false });
    }
    // The next add must throw — not silently produce a truncated archive.
    await expect(writer.addEntry('overflow.txt', payload)).rejects.toBeInstanceOf(
      OpenXmlNotImplementedError,
    );
  }, /* timeout */ 120_000);
});
