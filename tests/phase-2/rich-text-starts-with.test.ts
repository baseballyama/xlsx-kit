// Tests for the public `richTextStartsWith(rt, search, fromIndex?)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextStartsWith } from '../../src/cell';

describe('richTextStartsWith', () => {
  it('matches the prefix when search lies entirely within the first run', () => {
    const rt = makeRichText([{ text: 'hello world', font: { b: true } }]);
    expect(richTextStartsWith(rt, 'hello')).toBe(true);
    expect(richTextStartsWith(rt, '')).toBe(true);
  });

  it('matches the prefix across run boundaries', () => {
    const rt = makeRichText([
      { text: 'hel', font: { b: true } },
      { text: 'lo wor', font: { i: true } },
      { text: 'ld' },
    ]);
    expect(richTextStartsWith(rt, 'hello wo')).toBe(true);
    expect(richTextStartsWith(rt, 'hello world')).toBe(true);
  });

  it('returns false when the prefix does not match', () => {
    const rt = makeRichText([{ text: 'hello' }]);
    expect(richTextStartsWith(rt, 'world')).toBe(false);
    expect(richTextStartsWith(rt, 'helloX')).toBe(false);
  });

  it('honors fromIndex by matching at the offset position', () => {
    const rt = makeRichText([{ text: 'abc' }, { text: 'def' }]);
    expect(richTextStartsWith(rt, 'cde', 2)).toBe(true);
    expect(richTextStartsWith(rt, 'def', 3)).toBe(true);
    expect(richTextStartsWith(rt, 'abc', 1)).toBe(false);
  });
});
