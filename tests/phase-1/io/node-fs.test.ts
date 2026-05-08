// Node filesystem + Readable/Writable I/O helpers.
// Per docs/plan/03-foundations.md §1.1.

import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable, Writable } from 'node:stream';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { fromFile, fromFileSync, fromReadable, toFile, toWritable } from '../../../src/io/node-fs';
import { OpenXmlIoError } from '../../../src/utils/exceptions';

let scratch: string;
beforeAll(() => {
  scratch = mkdtempSync(join(tmpdir(), 'xlsx-craft-io-'));
});
afterAll(() => {
  rmSync(scratch, { recursive: true, force: true });
});

describe('fromFile', () => {
  it('reads bytes via toBytes()', async () => {
    const path = join(scratch, 'a.bin');
    writeFileSync(path, Buffer.from('hello fromFile', 'utf8'));
    const src = fromFile(path);
    const bytes = await src.toBytes();
    expect(new TextDecoder().decode(bytes)).toBe('hello fromFile');
  });

  it('streams via toStream() and yields the same bytes', async () => {
    const path = join(scratch, 'b.bin');
    const payload = new Uint8Array(64 * 1024);
    for (let i = 0; i < payload.byteLength; i++) payload[i] = i & 0xff;
    writeFileSync(path, payload);
    const src = fromFile(path);
    const stream = src.toStream?.();
    if (!stream) throw new Error('expected stream');
    const chunks: Uint8Array[] = [];
    const reader = stream.getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    let total = 0;
    for (const c of chunks) total += c.byteLength;
    const out = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      out.set(c, off);
      off += c.byteLength;
    }
    expect(out).toEqual(payload);
  });

  it('rejects empty paths with OpenXmlIoError', () => {
    expect(() => fromFile('')).toThrowError(OpenXmlIoError);
  });

  it('surfaces ENOENT through toBytes() as OpenXmlIoError', async () => {
    const src = fromFile(join(scratch, 'does-not-exist.bin'));
    await expect(src.toBytes()).rejects.toBeInstanceOf(OpenXmlIoError);
  });
});

describe('fromFileSync', () => {
  it('reads file synchronously', async () => {
    const path = join(scratch, 'sync.bin');
    writeFileSync(path, Buffer.from('sync hello', 'utf8'));
    const src = fromFileSync(path);
    const bytes = await src.toBytes();
    expect(new TextDecoder().decode(bytes)).toBe('sync hello');
  });

  it('throws OpenXmlIoError on missing file', () => {
    expect(() => fromFileSync(join(scratch, 'nope.bin'))).toThrowError(OpenXmlIoError);
  });
});

describe('toFile', () => {
  it('writes incoming chunks to disk and returns the path via result()', async () => {
    const path = join(scratch, 'out.bin');
    const sink = toFile(path);
    const w = sink.toBytes();
    w.write(new Uint8Array([1, 2, 3]));
    w.write(new Uint8Array([4, 5]));
    await w.finish();
    expect(sink.result()).toBe(path);
    expect(Array.from(readFileSync(path))).toEqual([1, 2, 3, 4, 5]);
  });

  it('rejects writes after finish()', async () => {
    const path = join(scratch, 'closed.bin');
    const sink = toFile(path);
    const w = sink.toBytes();
    w.write(new Uint8Array([1]));
    await w.finish();
    expect(() => w.write(new Uint8Array([2]))).toThrowError(OpenXmlIoError);
  });

  it('rejects empty paths', () => {
    expect(() => toFile('')).toThrowError(OpenXmlIoError);
  });
});

describe('fromReadable', () => {
  it('collects a Node Readable into a single Uint8Array', async () => {
    const r = Readable.from([Buffer.from('foo'), Buffer.from('bar'), Buffer.from('baz')]);
    const src = fromReadable(r);
    const bytes = await src.toBytes();
    expect(new TextDecoder().decode(bytes)).toBe('foobarbaz');
  });

  it('rejects non-Readable inputs', () => {
    // @ts-expect-error deliberately wrong
    expect(() => fromReadable('not a stream')).toThrowError(OpenXmlIoError);
  });
});

describe('toWritable', () => {
  it('forwards chunks to an underlying Writable', async () => {
    const collected: Uint8Array[] = [];
    const w = new Writable({
      write(chunk, _enc, cb) {
        collected.push(chunk instanceof Uint8Array ? new Uint8Array(chunk) : new Uint8Array(chunk));
        cb();
      },
    });
    const sink = toWritable(w);
    const sw = sink.toBytes();
    sw.write(new Uint8Array([1, 2]));
    sw.write(new Uint8Array([3, 4, 5]));
    await sw.finish();
    let total = 0;
    for (const c of collected) total += c.byteLength;
    expect(total).toBe(5);
    expect(sink.result()).toBe(w);
  });

  it('rejects non-Writable inputs', () => {
    // @ts-expect-error deliberately wrong
    expect(() => toWritable({})).toThrowError(OpenXmlIoError);
  });
});
