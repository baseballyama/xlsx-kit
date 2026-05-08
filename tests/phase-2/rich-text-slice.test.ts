// Tests for the public `sliceRichText(rt, start, end?)` substring helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextToString, sliceRichText } from '../../src/cell';

describe('sliceRichText', () => {
  it('slices within a single run, preserving font', () => {
    const rt = makeRichText([{ text: 'abcdef', font: { b: true } }]);
    const out = sliceRichText(rt, 1, 4);
    expect(out.length).toBe(1);
    expect(out[0]?.text).toBe('bcd');
    expect(out[0]?.font).toEqual({ b: true });
  });

  it('crosses run boundaries, keeping each run\'s font', () => {
    const rt = makeRichText([
      { text: 'hello', font: { b: true } },
      { text: ' world', font: { i: true } },
    ]);
    const out = sliceRichText(rt, 3, 8);
    expect(richTextToString(out)).toBe('lo wo');
    expect(out.length).toBe(2);
    expect(out[0]?.text).toBe('lo');
    expect(out[0]?.font).toEqual({ b: true });
    expect(out[1]?.text).toBe(' wo');
    expect(out[1]?.font).toEqual({ i: true });
  });

  it('treats negative indices like String.prototype.slice', () => {
    const rt = makeRichText([{ text: 'abcdef' }]);
    expect(richTextToString(sliceRichText(rt, -3))).toBe('def');
    expect(richTextToString(sliceRichText(rt, -4, -1))).toBe('cde');
  });

  it('with end omitted, slices through to the end', () => {
    const rt = makeRichText([{ text: 'ab' }, { text: 'cd' }]);
    expect(richTextToString(sliceRichText(rt, 1))).toBe('bcd');
  });

  it('returns an empty RichText for empty or reversed ranges', () => {
    const rt = makeRichText([{ text: 'abc' }]);
    expect(sliceRichText(rt, 2, 2).length).toBe(0);
    expect(sliceRichText(rt, 5, 8).length).toBe(0);
    expect(sliceRichText(rt, 4, 1).length).toBe(0);
  });
});
