// Tests for the public `findAllRichTextIndex(rt, search)` helper.

import { describe, expect, it } from 'vitest';
import { findAllRichTextIndex, makeRichText } from '../../src/cell';

describe('findAllRichTextIndex', () => {
  it('returns every non-overlapping occurrence within a single run', () => {
    const rt = makeRichText([{ text: 'banana' }]);
    expect(findAllRichTextIndex(rt, 'an')).toEqual([1, 3]);
    expect(findAllRichTextIndex(rt, 'a')).toEqual([1, 3, 5]);
  });

  it('returns occurrences spanning run boundaries', () => {
    const rt = makeRichText([
      { text: 'foo|', font: { b: true } },
      { text: 'foo|', font: { i: true } },
      { text: 'foo' },
    ]);
    expect(findAllRichTextIndex(rt, 'foo')).toEqual([0, 4, 8]);
  });

  it('returns an empty array when search is not found', () => {
    expect(findAllRichTextIndex(makeRichText([{ text: 'hello' }]), 'xyz')).toEqual([]);
    expect(findAllRichTextIndex(makeRichText([]), 'a')).toEqual([]);
  });

  it('returns an empty array for an empty search', () => {
    expect(findAllRichTextIndex(makeRichText([{ text: 'abc' }]), '')).toEqual([]);
  });
});
