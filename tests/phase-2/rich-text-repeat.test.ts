// Tests for the public `repeatRichText(rt, count)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, repeatRichText, richTextToString } from '../../src/cell';

describe('repeatRichText', () => {
  it('repeats multiple runs preserving fonts and order', () => {
    const rt = makeRichText([
      { text: 'AB', font: { b: true } },
      { text: 'cd', font: { i: true } },
    ]);
    const out = repeatRichText(rt, 3);
    expect(richTextToString(out)).toBe('ABcdABcdABcd');
    expect(out.length).toBe(6);
    for (let i = 0; i < 6; i += 2) {
      expect(out[i]?.text).toBe('AB');
      expect(out[i]?.font).toEqual({ b: true });
      expect(out[i + 1]?.text).toBe('cd');
      expect(out[i + 1]?.font).toEqual({ i: true });
    }
  });

  it('returns an empty RichText when count is 0', () => {
    const rt = makeRichText([{ text: 'x', font: { b: true } }]);
    const out = repeatRichText(rt, 0);
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });

  it('returns the input unchanged when count is 1', () => {
    const rt = makeRichText([{ text: 'x' }]);
    expect(repeatRichText(rt, 1)).toBe(rt);
  });

  it('returns an empty RichText when rt is empty regardless of count', () => {
    const out = repeatRichText(makeRichText([]), 5);
    expect(out.length).toBe(0);
  });

  it('throws RangeError on a negative count', () => {
    const rt = makeRichText([{ text: 'x' }]);
    expect(() => repeatRichText(rt, -1)).toThrow(RangeError);
    expect(() => repeatRichText(rt, Number.NaN)).toThrow(RangeError);
    expect(() => repeatRichText(rt, Number.POSITIVE_INFINITY)).toThrow(RangeError);
  });
});
