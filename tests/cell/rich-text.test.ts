import { describe, expect, it } from 'vitest';
import { makeRichText, makeTextRun, richTextToString } from '../../src/cell/rich-text';
import { makeColor } from '../../src/styles/colors';

describe('makeTextRun', () => {
  it('builds a frozen run with text-only', () => {
    const r = makeTextRun('hello');
    expect(Object.isFrozen(r)).toBe(true);
    expect(r.text).toBe('hello');
    expect(r.font).toBeUndefined();
  });

  it('attaches an InlineFont when supplied', () => {
    const font = { name: 'Calibri', sz: 11, b: true, color: makeColor({ rgb: 'FF0000' }) };
    const r = makeTextRun('bold!', font);
    expect(r.font).toBe(font);
  });

  it('rejects non-string text', () => {
    // biome-ignore lint/suspicious/noExplicitAny: deliberate bad input
    expect(() => makeTextRun(42 as any)).toThrowError(TypeError);
  });
});

describe('makeRichText', () => {
  it('freezes the array and each run', () => {
    const rt = makeRichText([{ text: 'a' }, { text: 'b', font: { b: true } }]);
    expect(Object.isFrozen(rt)).toBe(true);
    expect(rt[0] && Object.isFrozen(rt[0])).toBe(true);
    expect(rt[1]?.font?.b).toBe(true);
  });

  it('passes through frozen TextRun inputs', () => {
    const r = makeTextRun('x', { i: true });
    const rt = makeRichText([r]);
    expect(rt[0]).toBe(r);
  });
});

describe('richTextToString', () => {
  it('concatenates all runs as plain text', () => {
    const rt = makeRichText([makeTextRun('hello '), makeTextRun('world', { b: true }), makeTextRun('!')]);
    expect(richTextToString(rt)).toBe('hello world!');
  });

  it('returns "" for an empty rich-text value', () => {
    expect(richTextToString(makeRichText([]))).toBe('');
  });
});
