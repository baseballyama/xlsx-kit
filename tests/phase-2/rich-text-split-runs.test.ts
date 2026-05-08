// Tests for the public `splitRichTextRuns(rt)` per-character split helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextToString, splitRichTextRuns } from '../../src/cell';

describe('splitRichTextRuns', () => {
  it('splits a single ASCII run into one run per character, preserving font', () => {
    const rt = makeRichText([{ text: 'abc', font: { b: true } }]);
    const out = splitRichTextRuns(rt);
    expect(out.length).toBe(3);
    expect(out.map((r) => r.text)).toEqual(['a', 'b', 'c']);
    for (const r of out) expect(r.font).toEqual({ b: true });
    expect(Object.isFrozen(out)).toBe(true);
  });

  it('splits across multiple runs, each carrying its parent run font', () => {
    const rt = makeRichText([
      { text: 'ab' },
      { text: 'cd', font: { i: true } },
    ]);
    const out = splitRichTextRuns(rt);
    expect(out.length).toBe(4);
    expect(richTextToString(out)).toBe('abcd');
    expect(out[0]?.font).toBeUndefined();
    expect(out[1]?.font).toBeUndefined();
    expect(out[2]?.font).toEqual({ i: true });
    expect(out[3]?.font).toEqual({ i: true });
  });

  it('drops empty-text runs without producing zero-length outputs', () => {
    const rt = makeRichText([{ text: '' }, { text: 'x' }, { text: '' }]);
    const out = splitRichTextRuns(rt);
    expect(out.length).toBe(1);
    expect(out[0]?.text).toBe('x');
  });

  it('returns an empty RichText when input is empty', () => {
    const out = splitRichTextRuns(makeRichText([]));
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });
});
