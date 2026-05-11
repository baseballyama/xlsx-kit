import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../../src/io/node';
import { OpenXmlIoError } from '../../../src/utils/exceptions';

describe('fromBuffer', () => {
  it('exposes a Uint8Array source from a Node Buffer', async () => {
    const src = fromBuffer(Buffer.from('hello world', 'utf8'));
    const bytes = await src.toBytes();
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(new TextDecoder().decode(bytes)).toBe('hello world');
  });

  it('exposes a Uint8Array source from a Uint8Array', async () => {
    const u8 = new Uint8Array([1, 2, 3]);
    const src = fromBuffer(u8);
    expect(await src.toBytes()).toEqual(u8);
  });

  it('handles 0-byte input without throwing', async () => {
    const src = fromBuffer(new Uint8Array(0));
    const bytes = await src.toBytes();
    expect(bytes.byteLength).toBe(0);
  });

  it('exposes a stream that yields the same payload', async () => {
    const u8 = new Uint8Array([9, 8, 7, 6]);
    const src = fromBuffer(u8);
    const stream = src.toStream?.();
    if (!stream) throw new Error('expected toStream() to be defined');
    const reader = stream.getReader();
    const chunks: number[] = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(...value);
    }
    expect(chunks).toEqual([9, 8, 7, 6]);
  });

  it('rejects non-byte inputs with OpenXmlIoError', () => {
    // biome-ignore lint/suspicious/noExplicitAny: deliberately exercising the error path
    expect(() => fromBuffer('not bytes' as any)).toThrowError(OpenXmlIoError);
  });
});

describe('toBuffer', () => {
  it('accumulates writes and yields them via finish()', async () => {
    const sink = toBuffer();
    const w = sink.toBytes();
    w.write(new Uint8Array([1, 2]));
    w.write(new Uint8Array([3, 4, 5]));
    const bytes = await w.finish();
    expect(Array.from(bytes)).toEqual([1, 2, 3, 4, 5]);
  });

  it('exposes the same payload via the result() helper as a Buffer', async () => {
    const sink = toBuffer();
    const w = sink.toBytes();
    w.write(new Uint8Array([10, 20, 30]));
    await w.finish();
    const buf = sink.result();
    expect(Buffer.isBuffer(buf)).toBe(true);
    expect(Array.from(buf)).toEqual([10, 20, 30]);
  });

  it('result() is callable without an explicit finish() and returns the buffered bytes', () => {
    const sink = toBuffer();
    const w = sink.toBytes();
    w.write(new Uint8Array([42]));
    const buf = sink.result();
    expect(Array.from(buf)).toEqual([42]);
  });

  it('throws OpenXmlIoError on writing after finish', async () => {
    const sink = toBuffer();
    const w = sink.toBytes();
    w.write(new Uint8Array([1]));
    await w.finish();
    expect(() => w.write(new Uint8Array([2]))).toThrowError(OpenXmlIoError);
  });

  it('throws OpenXmlIoError when a non-Uint8Array chunk is appended', () => {
    const sink = toBuffer();
    const w = sink.toBytes();
    // biome-ignore lint/suspicious/noExplicitAny: deliberately exercising the error path
    expect(() => w.write([1, 2, 3] as any)).toThrowError(OpenXmlIoError);
  });

  it('handles 0-byte payloads cleanly', async () => {
    const sink = toBuffer();
    await sink.toBytes().finish();
    const buf = sink.result();
    expect(buf.byteLength).toBe(0);
  });
});
