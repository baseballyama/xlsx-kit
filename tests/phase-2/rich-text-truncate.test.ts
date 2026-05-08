// Tests for the public `truncateRichText(rt, maxLength, ellipsis?)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextToString, truncateRichText } from '../../src/cell';

describe('truncateRichText', () => {
  it('returns the input unchanged when already short enough', () => {
    const rt = makeRichText([{ text: 'hello', font: { b: true } }]);
    expect(truncateRichText(rt, 10)).toBe(rt);
    expect(truncateRichText(rt, 5)).toBe(rt);
  });

  it('hard-truncates without ellipsis when too long', () => {
    const rt = makeRichText([
      { text: 'hello ', font: { b: true } },
      { text: 'world', font: { i: true } },
    ]);
    const out = truncateRichText(rt, 7);
    expect(richTextToString(out)).toBe('hello w');
    expect(out[0]?.font).toEqual({ b: true });
    expect(out[1]?.font).toEqual({ i: true });
  });

  it('appends an ellipsis as a font-less trailing run', () => {
    const rt = makeRichText([{ text: 'hello world', font: { b: true } }]);
    const out = truncateRichText(rt, 8, '...');
    expect(richTextToString(out)).toBe('hello...');
    expect(out[0]?.font).toEqual({ b: true });
    expect(out[out.length - 1]?.font).toBeUndefined();
  });

  it('returns an empty RichText when maxLength <= 0', () => {
    const rt = makeRichText([{ text: 'abc' }]);
    expect(truncateRichText(rt, 0).length).toBe(0);
    expect(truncateRichText(rt, -5, '...').length).toBe(0);
  });

  it('hard-truncates the ellipsis itself when it is wider than maxLength', () => {
    const rt = makeRichText([{ text: 'abcdef' }]);
    const out = truncateRichText(rt, 2, '...');
    expect(richTextToString(out)).toBe('..');
    expect(out[0]?.font).toBeUndefined();
  });
});
