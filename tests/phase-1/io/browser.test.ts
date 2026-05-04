import { describe, expect, it } from 'vitest';
import {
  fromArrayBuffer,
  fromBlob,
  fromFile,
  fromResponse,
  fromStream,
  toArrayBuffer,
  toBlob,
} from '../../../src/io/browser';
import { OpenXmlIoError } from '../../../src/utils/exceptions';

describe('fromBlob', () => {
  it('reads bytes via toBytes()', async () => {
    const blob = new Blob([new Uint8Array([1, 2, 3])]);
    const src = fromBlob(blob);
    const bytes = await src.toBytes();
    expect(Array.from(bytes)).toEqual([1, 2, 3]);
  });

  it('exposes a stream for sequential read', async () => {
    const blob = new Blob([new Uint8Array([7, 8, 9])]);
    const src = fromBlob(blob);
    const stream = src.toStream?.();
    if (!stream) throw new Error('expected toStream() to be defined');
    const reader = stream.getReader();
    const out: number[] = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      out.push(...value);
    }
    expect(out).toEqual([7, 8, 9]);
  });

  it('handles 0-byte blobs', async () => {
    const src = fromBlob(new Blob([]));
    expect((await src.toBytes()).byteLength).toBe(0);
  });

  it('rejects non-Blob input with OpenXmlIoError', () => {
    // biome-ignore lint/suspicious/noExplicitAny: deliberately exercising the error path
    expect(() => fromBlob('not a blob' as any)).toThrowError(OpenXmlIoError);
  });

  it('fromFile is a synonym for fromBlob (File extends Blob)', () => {
    expect(fromFile).toBe(fromBlob);
  });
});

describe('fromArrayBuffer', () => {
  it('wraps an ArrayBuffer', async () => {
    const ab = new ArrayBuffer(3);
    new Uint8Array(ab).set([1, 2, 3]);
    const src = fromArrayBuffer(ab);
    expect(Array.from(await src.toBytes())).toEqual([1, 2, 3]);
  });

  it('wraps a Uint8Array (no copy semantics)', async () => {
    const u8 = new Uint8Array([10, 20]);
    const src = fromArrayBuffer(u8);
    expect((await src.toBytes()) === u8).toBe(true);
  });

  it('rejects other inputs with OpenXmlIoError', () => {
    // biome-ignore lint/suspicious/noExplicitAny: deliberately exercising the error path
    expect(() => fromArrayBuffer({} as any)).toThrowError(OpenXmlIoError);
  });
});

describe('toBlob', () => {
  it('accumulates writes and yields them as a Blob with the chosen MIME', async () => {
    const sink = toBlob('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const w = sink.toBytes();
    w.write(new Uint8Array([1, 2]));
    w.write(new Uint8Array([3]));
    await w.finish();
    const blob = sink.result();
    expect(blob).toBeInstanceOf(Blob);
    expect(blob.type).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(Array.from(new Uint8Array(await blob.arrayBuffer()))).toEqual([1, 2, 3]);
  });

  it('defaults to application/octet-stream when no MIME is supplied', () => {
    const sink = toBlob();
    const blob = sink.result();
    expect(blob.type).toBe('application/octet-stream');
  });

  it('throws OpenXmlIoError when writing after finish', async () => {
    const sink = toBlob();
    const w = sink.toBytes();
    await w.finish();
    expect(() => w.write(new Uint8Array([1]))).toThrowError(OpenXmlIoError);
  });
});

describe('fromResponse', () => {
  it('reads bytes via toBytes() from a JSON-style Response', async () => {
    const payload = new Uint8Array([1, 2, 3, 4]);
    const res = new Response(payload);
    const src = fromResponse(res);
    expect(Array.from(await src.toBytes())).toEqual([1, 2, 3, 4]);
  });

  it('exposes the body stream via toStream()', async () => {
    const payload = new Uint8Array([9, 8, 7]);
    const res = new Response(payload);
    const src = fromResponse(res);
    const stream = src.toStream?.();
    if (!stream) throw new Error('expected toStream() to be defined');
    const reader = stream.getReader();
    const out: number[] = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      out.push(...value);
    }
    expect(out).toEqual([9, 8, 7]);
  });

  it('handles bodyless responses with an empty stream', async () => {
    const res = new Response(null, { status: 204 });
    const src = fromResponse(res);
    const stream = src.toStream?.();
    if (!stream) throw new Error('expected toStream() to be defined');
    const reader = stream.getReader();
    const { done } = await reader.read();
    expect(done).toBe(true);
  });

  it('rejects non-Response inputs with OpenXmlIoError', () => {
    // biome-ignore lint/suspicious/noExplicitAny: deliberately exercising the error path
    expect(() => fromResponse({} as any)).toThrowError(OpenXmlIoError);
  });
});

describe('fromStream', () => {
  const streamOf = (chunks: Uint8Array[]): ReadableStream<Uint8Array> =>
    new ReadableStream<Uint8Array>({
      start(controller) {
        for (const c of chunks) controller.enqueue(c);
        controller.close();
      },
    });

  it('drains a Web ReadableStream into Uint8Array via toBytes()', async () => {
    const src = fromStream(streamOf([new Uint8Array([1, 2]), new Uint8Array([3, 4, 5])]));
    expect(Array.from(await src.toBytes())).toEqual([1, 2, 3, 4, 5]);
  });

  it('returns the same stream via toStream()', async () => {
    const original = streamOf([new Uint8Array([42])]);
    const src = fromStream(original);
    const back = src.toStream?.();
    expect(back).toBe(original);
  });

  it('throws when toBytes is called after toStream', () => {
    const src = fromStream(streamOf([new Uint8Array([1])]));
    src.toStream?.();
    return expect(src.toBytes()).rejects.toBeInstanceOf(OpenXmlIoError);
  });

  it('throws when toStream is called after toBytes', async () => {
    const src = fromStream(streamOf([new Uint8Array([1])]));
    await src.toBytes();
    expect(() => src.toStream?.()).toThrowError(OpenXmlIoError);
  });

  it('rejects non-ReadableStream inputs', () => {
    // biome-ignore lint/suspicious/noExplicitAny: deliberately exercising the error path
    expect(() => fromStream('not a stream' as any)).toThrowError(OpenXmlIoError);
  });
});

describe('toArrayBuffer', () => {
  it('accumulates writes and yields them as an ArrayBuffer', async () => {
    const sink = toArrayBuffer();
    const w = sink.toBytes();
    w.write(new Uint8Array([1, 2]));
    w.write(new Uint8Array([3, 4]));
    await w.finish();
    const out = sink.result();
    expect(out).toBeInstanceOf(ArrayBuffer);
    expect(Array.from(new Uint8Array(out))).toEqual([1, 2, 3, 4]);
  });

  it('returns an exactly-sized ArrayBuffer (no oversized backing buffer)', async () => {
    const sink = toArrayBuffer();
    const w = sink.toBytes();
    w.write(new Uint8Array([7]));
    await w.finish();
    const out = sink.result();
    expect(out.byteLength).toBe(1);
  });
});
