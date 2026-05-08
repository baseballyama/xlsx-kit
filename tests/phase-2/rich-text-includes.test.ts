// Tests for the public `richTextIncludes(rt, search, fromIndex?)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextIncludes } from '../../src/cell';

describe('richTextIncludes', () => {
  it('returns true when search is contained within a single run', () => {
    const rt = makeRichText([{ text: 'hello world', font: { b: true } }]);
    expect(richTextIncludes(rt, 'world')).toBe(true);
  });

  it('returns true when search spans run boundaries', () => {
    const rt = makeRichText([
      { text: 'hel', font: { b: true } },
      { text: 'lo wor', font: { i: true } },
      { text: 'ld' },
    ]);
    expect(richTextIncludes(rt, 'lo wo')).toBe(true);
    expect(richTextIncludes(rt, 'hello world')).toBe(true);
  });

  it('returns false when search is absent', () => {
    const rt = makeRichText([{ text: 'hello' }]);
    expect(richTextIncludes(rt, 'xyz')).toBe(false);
    expect(richTextIncludes(rt, 'hello', 1)).toBe(false);
  });

  it('returns true for an empty search regardless of content', () => {
    expect(richTextIncludes(makeRichText([{ text: 'abc' }]), '')).toBe(true);
    expect(richTextIncludes(makeRichText([]), '')).toBe(true);
  });
});
