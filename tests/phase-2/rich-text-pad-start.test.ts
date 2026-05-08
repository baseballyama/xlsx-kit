// Tests for the public `padStartRichText(rt, targetLength, padString?)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, padStartRichText, richTextToString } from '../../src/cell';

describe('padStartRichText', () => {
  it('pads with default space until targetLength is reached, preserving existing run fonts', () => {
    const rt = makeRichText([{ text: 'hi', font: { b: true } }]);
    const out = padStartRichText(rt, 5);
    expect(richTextToString(out)).toBe('   hi');
    expect(out[0]?.font).toBeUndefined();
    expect(out[out.length - 1]?.text).toBe('hi');
    expect(out[out.length - 1]?.font).toEqual({ b: true });
  });

  it('handles a custom padString that does not divide the gap evenly', () => {
    const rt = makeRichText([{ text: 'X' }]);
    const out = padStartRichText(rt, 6, 'ab');
    expect(richTextToString(out)).toBe('ababaX');
  });

  it('returns the input unchanged when targetLength is not greater than the current length', () => {
    const rt = makeRichText([{ text: 'hello' }]);
    expect(padStartRichText(rt, 5)).toBe(rt);
    expect(padStartRichText(rt, 3)).toBe(rt);
  });

  it('returns the input unchanged when padString is empty', () => {
    const rt = makeRichText([{ text: 'abc' }]);
    expect(padStartRichText(rt, 10, '')).toBe(rt);
  });
});
