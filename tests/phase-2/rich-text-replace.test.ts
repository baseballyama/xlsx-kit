// Tests for the public `replaceRichText(rt, start, end, replacement)` helper.

import { describe, expect, it } from 'vitest';
import {
  makeRichText,
  replaceRichText,
  richText,
  richTextLength,
  richTextToString,
} from '../../src/cell';

describe('replaceRichText', () => {
  it('replaces a substring with a font-less string', () => {
    const rt = makeRichText([{ text: 'hello world' }]);
    const out = replaceRichText(rt, 6, 11, 'there');
    expect(richTextToString(out)).toBe('hello there');
  });

  it('preserves replacement run fonts', () => {
    const rt = makeRichText([{ text: 'a__c' }]);
    const replacement = richText('B', { b: true });
    const out = replaceRichText(rt, 1, 3, replacement);
    expect(richTextToString(out)).toBe('aBc');
    const boldRun = out.find((r) => r.text === 'B');
    expect(boldRun?.font).toEqual({ b: true });
  });

  it('deletes the range when replacement is the empty string', () => {
    const rt = makeRichText([{ text: 'abcdef' }]);
    const out = replaceRichText(rt, 2, 4, '');
    expect(richTextToString(out)).toBe('abef');
    expect(richTextLength(out)).toBe(4);
  });

  it('handles negative indices like String.prototype.slice', () => {
    const rt = makeRichText([{ text: 'abcdef' }]);
    const out = replaceRichText(rt, -3, -1, 'XY');
    expect(richTextToString(out)).toBe('abcXYf');
  });

  it('replaces through the end when end >= length', () => {
    const rt = makeRichText([{ text: 'hello' }]);
    const out = replaceRichText(rt, 3, 99, '!!');
    expect(richTextToString(out)).toBe('hel!!');
  });
});
