// Tests for the public `concatRichText(...parts)` flatten helper.

import { describe, expect, it } from 'vitest';
import { concatRichText, makeTextRun, richText, richTextToString } from '../../src/cell';

describe('concatRichText', () => {
  it('flattens multiple strings into one font-less run per string', () => {
    const out = concatRichText('a', 'b', 'c');
    expect(out.length).toBe(3);
    expect(richTextToString(out)).toBe('abc');
    for (const r of out) expect(r.font).toBeUndefined();
  });

  it('spreads RichText parts in order', () => {
    const a = richText('hello', { b: true });
    const b = richText('world', { i: true });
    const out = concatRichText(a, b);
    expect(out.length).toBe(2);
    expect(out[0]?.text).toBe('hello');
    expect(out[0]?.font).toEqual({ b: true });
    expect(out[1]?.text).toBe('world');
    expect(out[1]?.font).toEqual({ i: true });
  });

  it('mixes string + TextRun + RichText in one call', () => {
    const out = concatRichText('start ', makeTextRun('mid', { b: true }), richText(' end'));
    expect(out.length).toBe(3);
    expect(richTextToString(out)).toBe('start mid end');
    expect(out[1]?.font).toEqual({ b: true });
    expect(out[0]?.font).toBeUndefined();
    expect(out[2]?.font).toBeUndefined();
  });

  it('returns an empty frozen RichText when called with no args', () => {
    const out = concatRichText();
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });
});
