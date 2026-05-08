// Tests for the public `getRichTextCharAt(rt, index)` helper.

import { describe, expect, it } from 'vitest';
import { getRichTextCharAt, makeRichText } from '../../src/cell';

describe('getRichTextCharAt', () => {
  it('returns characters by index within a single run', () => {
    const rt = makeRichText([{ text: 'hello', font: { b: true } }]);
    expect(getRichTextCharAt(rt, 0)).toBe('h');
    expect(getRichTextCharAt(rt, 4)).toBe('o');
  });

  it('returns the right character across run boundaries', () => {
    const rt = makeRichText([
      { text: 'AB', font: { b: true } },
      { text: 'cd', font: { i: true } },
      { text: 'EF' },
    ]);
    expect(getRichTextCharAt(rt, 0)).toBe('A');
    expect(getRichTextCharAt(rt, 2)).toBe('c');
    expect(getRichTextCharAt(rt, 4)).toBe('E');
    expect(getRichTextCharAt(rt, 5)).toBe('F');
  });

  it('returns an empty string for out-of-range indices', () => {
    const rt = makeRichText([{ text: 'abc' }]);
    expect(getRichTextCharAt(rt, -1)).toBe('');
    expect(getRichTextCharAt(rt, 3)).toBe('');
    expect(getRichTextCharAt(rt, 100)).toBe('');
    expect(getRichTextCharAt(makeRichText([]), 0)).toBe('');
  });
});
