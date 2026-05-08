// Tests for the public `countRichTextOccurrences(rt, search)` helper.

import { describe, expect, it } from 'vitest';
import { countRichTextOccurrences, makeRichText } from '../../src/cell';

describe('countRichTextOccurrences', () => {
  it('counts multiple non-overlapping occurrences within a single run', () => {
    const rt = makeRichText([{ text: 'banana' }]);
    expect(countRichTextOccurrences(rt, 'an')).toBe(2);
    expect(countRichTextOccurrences(rt, 'a')).toBe(3);
  });

  it('counts occurrences spanning run boundaries', () => {
    const rt = makeRichText([
      { text: 'foo|', font: { b: true } },
      { text: 'foo|', font: { i: true } },
      { text: 'foo' },
    ]);
    expect(countRichTextOccurrences(rt, 'foo')).toBe(3);
    expect(countRichTextOccurrences(rt, 'oo|f')).toBe(2);
  });

  it('returns 0 when search is not present', () => {
    expect(countRichTextOccurrences(makeRichText([{ text: 'hello' }]), 'xyz')).toBe(0);
    expect(countRichTextOccurrences(makeRichText([]), 'a')).toBe(0);
  });

  it('returns 0 for an empty search', () => {
    expect(countRichTextOccurrences(makeRichText([{ text: 'abc' }]), '')).toBe(0);
  });
});
