// Tests for the public `richTextEndsWith(rt, search, endIndex?)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextEndsWith } from '../../src/cell';

describe('richTextEndsWith', () => {
  it('matches the suffix when search lies entirely within the last run', () => {
    const rt = makeRichText([{ text: 'hello world', font: { b: true } }]);
    expect(richTextEndsWith(rt, 'world')).toBe(true);
    expect(richTextEndsWith(rt, '')).toBe(true);
  });

  it('matches the suffix across run boundaries', () => {
    const rt = makeRichText([
      { text: 'hel', font: { b: true } },
      { text: 'lo wor', font: { i: true } },
      { text: 'ld' },
    ]);
    expect(richTextEndsWith(rt, 'world')).toBe(true);
    expect(richTextEndsWith(rt, 'hello world')).toBe(true);
  });

  it('returns false when the suffix does not match', () => {
    const rt = makeRichText([{ text: 'hello' }]);
    expect(richTextEndsWith(rt, 'world')).toBe(false);
    expect(richTextEndsWith(rt, 'Xhello')).toBe(false);
  });

  it('honors endIndex by truncating the considered suffix', () => {
    const rt = makeRichText([{ text: 'abc' }, { text: 'def' }]);
    expect(richTextEndsWith(rt, 'abc', 3)).toBe(true);
    expect(richTextEndsWith(rt, 'def', 6)).toBe(true);
    expect(richTextEndsWith(rt, 'def', 5)).toBe(false);
  });
});
