// Tests for the public `padEndRichText(rt, targetLength, padString?)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, padEndRichText, richTextToString } from '../../src/cell';

describe('padEndRichText', () => {
  it('pads with default space until targetLength is reached, preserving existing run fonts', () => {
    const rt = makeRichText([{ text: 'hi', font: { b: true } }]);
    const out = padEndRichText(rt, 5);
    expect(richTextToString(out)).toBe('hi   ');
    expect(out[0]?.text).toBe('hi');
    expect(out[0]?.font).toEqual({ b: true });
    expect(out[1]?.font).toBeUndefined();
  });

  it('handles a custom padString that does not divide the gap evenly', () => {
    const rt = makeRichText([{ text: 'X' }]);
    const out = padEndRichText(rt, 6, 'ab');
    expect(richTextToString(out)).toBe('Xababa');
  });

  it('returns the input unchanged when targetLength is not greater than the current length', () => {
    const rt = makeRichText([{ text: 'hello' }]);
    expect(padEndRichText(rt, 5)).toBe(rt);
    expect(padEndRichText(rt, 3)).toBe(rt);
  });

  it('returns the input unchanged when padString is empty', () => {
    const rt = makeRichText([{ text: 'abc' }]);
    expect(padEndRichText(rt, 10, '')).toBe(rt);
  });
});
