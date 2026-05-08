// Tests for the public `richTextWords(rt)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextToString, richTextWords } from '../../src/cell';

describe('richTextWords', () => {
  it('splits a single run into words while preserving font', () => {
    const rt = makeRichText([{ text: '  hello  world\t', font: { b: true } }]);
    const words = richTextWords(rt);
    expect(words.map(richTextToString)).toEqual(['hello', 'world']);
    for (const w of words) expect(w[0]?.font).toEqual({ b: true });
  });

  it('preserves per-run fonts when a word spans run boundaries', () => {
    const rt = makeRichText([
      { text: 'hel', font: { b: true } },
      { text: 'lo wo', font: { i: true } },
      { text: 'rld' },
    ]);
    const words = richTextWords(rt);
    expect(words.map(richTextToString)).toEqual(['hello', 'world']);
    expect(words[0]?.[0]?.font).toEqual({ b: true });
    expect(words[0]?.[1]?.font).toEqual({ i: true });
    expect(words[1]?.[0]?.font).toEqual({ i: true });
    expect(words[1]?.[1]?.font).toBeUndefined();
  });

  it('returns an empty array when the input is empty or all whitespace', () => {
    expect(richTextWords(makeRichText([]))).toEqual([]);
    expect(richTextWords(makeRichText([{ text: '  \t\r\n  ' }]))).toEqual([]);
  });

  it('drops consecutive whitespace and yields only non-empty word segments', () => {
    const rt = makeRichText([{ text: 'a  b\t\tc' }]);
    const words = richTextWords(rt);
    expect(words.map(richTextToString)).toEqual(['a', 'b', 'c']);
  });
});
