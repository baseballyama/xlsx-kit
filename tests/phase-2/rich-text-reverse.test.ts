// Tests for the public `reverseRichText(rt)` helper.

import { describe, expect, it } from 'vitest';
import {
  makeRichText,
  reverseRichText,
  richTextToString,
  splitRichTextRuns,
} from '../../src/cell';

describe('reverseRichText', () => {
  it('reverses a single run, preserving font', () => {
    const rt = makeRichText([{ text: 'abc', font: { b: true } }]);
    const out = reverseRichText(rt);
    expect(out.length).toBe(1);
    expect(out[0]?.text).toBe('cba');
    expect(out[0]?.font).toEqual({ b: true });
  });

  it('reverses both run order and each run text', () => {
    const rt = makeRichText([
      { text: 'AB', font: { b: true } },
      { text: 'CD', font: { i: true } },
    ]);
    const out = reverseRichText(rt);
    expect(richTextToString(out)).toBe('DCBA');
    expect(out.length).toBe(2);
    expect(out[0]?.text).toBe('DC');
    expect(out[0]?.font).toEqual({ i: true });
    expect(out[1]?.text).toBe('BA');
    expect(out[1]?.font).toEqual({ b: true });
  });

  it('returns an empty RichText when input is empty', () => {
    const out = reverseRichText(makeRichText([]));
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });

  it('produces the same concatenated text as reversing per-character splits', () => {
    const rt = makeRichText([
      { text: 'hello', font: { b: true } },
      { text: ' world' },
    ]);
    const direct = richTextToString(reverseRichText(rt));
    const viaSplit = richTextToString(reverseRichText(splitRichTextRuns(rt)));
    expect(direct).toBe(viaSplit);
    expect(direct).toBe('dlrow olleh');
  });
});
