// Tests for the public `findLastRichTextIndex(rt, search, fromIndex?)` helper.

import { describe, expect, it } from 'vitest';
import { findLastRichTextIndex, makeRichText } from '../../src/cell';

describe('findLastRichTextIndex', () => {
  it('finds the last occurrence within a single run', () => {
    const rt = makeRichText([{ text: 'banana' }]);
    expect(findLastRichTextIndex(rt, 'an')).toBe(3);
    expect(findLastRichTextIndex(rt, 'a')).toBe(5);
  });

  it('finds the last occurrence spanning run boundaries', () => {
    const rt = makeRichText([
      { text: 'foo|', font: { b: true } },
      { text: 'bar|', font: { i: true } },
      { text: 'foo' },
    ]);
    expect(findLastRichTextIndex(rt, 'foo')).toBe(8);
    expect(findLastRichTextIndex(rt, '|')).toBe(7);
  });

  it('honors fromIndex by limiting matches to positions at or before it', () => {
    const rt = makeRichText([{ text: 'banana' }]);
    expect(findLastRichTextIndex(rt, 'an', 2)).toBe(1);
    expect(findLastRichTextIndex(rt, 'an', 0)).toBe(-1);
  });

  it('returns -1 when search is not found', () => {
    const rt = makeRichText([{ text: 'hello' }]);
    expect(findLastRichTextIndex(rt, 'xyz')).toBe(-1);
    expect(findLastRichTextIndex(makeRichText([]), 'a')).toBe(-1);
  });
});
