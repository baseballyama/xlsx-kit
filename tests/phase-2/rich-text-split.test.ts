// Tests for the public `splitRichText(rt, separator, limit?)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextToString, splitRichText } from '../../src/cell';

describe('splitRichText', () => {
  it('splits a single run by a separator while preserving font', () => {
    const rt = makeRichText([{ text: 'a,b,c', font: { b: true } }]);
    const parts = splitRichText(rt, ',');
    expect(parts.map(richTextToString)).toEqual(['a', 'b', 'c']);
    for (const p of parts) {
      expect(p[0]?.font).toEqual({ b: true });
    }
  });

  it('splits across run boundaries, preserving each part runs fonts', () => {
    const rt = makeRichText([
      { text: 'foo|', font: { b: true } },
      { text: 'ba', font: { i: true } },
      { text: 'r|baz' },
    ]);
    const parts = splitRichText(rt, '|');
    expect(parts.map(richTextToString)).toEqual(['foo', 'bar', 'baz']);
    expect(parts[0]?.[0]?.font).toEqual({ b: true });
    expect(parts[1]?.[0]?.font).toEqual({ i: true });
    expect(parts[1]?.[1]?.font).toBeUndefined();
  });

  it('honors limit by truncating the resulting array (no remainder appended)', () => {
    const rt = makeRichText([{ text: 'a,b,c,d' }]);
    const parts = splitRichText(rt, ',', 2);
    expect(parts.map(richTextToString)).toEqual(['a', 'b']);
    expect(splitRichText(rt, ',', 0)).toEqual([]);
  });

  it('returns [rt] when separator is not found', () => {
    const rt = makeRichText([{ text: 'hello' }]);
    const parts = splitRichText(rt, 'xyz');
    expect(parts.length).toBe(1);
    const first = parts[0];
    expect(first && richTextToString(first)).toBe('hello');
  });

  it('splits per code unit when separator is empty', () => {
    const rt = makeRichText([{ text: 'AB', font: { b: true } }, { text: 'cd' }]);
    const parts = splitRichText(rt, '');
    expect(parts.map(richTextToString)).toEqual(['A', 'B', 'c', 'd']);
    expect(parts[0]?.[0]?.font).toEqual({ b: true });
    expect(parts[2]?.[0]?.font).toBeUndefined();
  });
});
