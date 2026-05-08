// Tests for the public `richTextLength` total-char-count helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richText, richTextLength } from '../../src/cell';

describe('richTextLength', () => {
  it('returns 0 for an empty RichText', () => {
    expect(richTextLength(makeRichText([]))).toBe(0);
  });

  it('returns the run text length for a single run', () => {
    expect(richTextLength(richText('hello'))).toBe(5);
    expect(richTextLength(richText(''))).toBe(0);
  });

  it('sums the lengths across multiple runs', () => {
    const rt = makeRichText([
      { text: 'hello ' },
      { text: 'world', font: { b: true } },
      { text: '!' },
    ]);
    expect(richTextLength(rt)).toBe(12);
  });
});
