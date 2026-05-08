// Tests for the public `findRichTextIndex(rt, search, fromIndex?)` helper.

import { describe, expect, it } from 'vitest';
import { findRichTextIndex, makeRichText } from '../../src/cell';

describe('findRichTextIndex', () => {
  it('finds a substring within a single run', () => {
    const rt = makeRichText([{ text: 'hello world' }]);
    expect(findRichTextIndex(rt, 'world')).toBe(6);
    expect(findRichTextIndex(rt, 'hello')).toBe(0);
  });

  it('finds a substring that spans multiple runs', () => {
    const rt = makeRichText([
      { text: 'hel', font: { b: true } },
      { text: 'lo wor', font: { i: true } },
      { text: 'ld' },
    ]);
    expect(findRichTextIndex(rt, 'lo wo')).toBe(3);
    expect(findRichTextIndex(rt, 'orld')).toBe(7);
  });

  it('uses fromIndex to find the second occurrence', () => {
    const rt = makeRichText([{ text: 'abcabc' }]);
    expect(findRichTextIndex(rt, 'abc')).toBe(0);
    expect(findRichTextIndex(rt, 'abc', 1)).toBe(3);
  });

  it('returns -1 when not found', () => {
    const rt = makeRichText([{ text: 'hello' }]);
    expect(findRichTextIndex(rt, 'world')).toBe(-1);
  });
});
