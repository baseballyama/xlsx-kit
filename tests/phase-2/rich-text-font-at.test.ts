// Tests for the public `getRichTextFontAt(rt, index)` helper.

import { describe, expect, it } from 'vitest';
import { getRichTextFontAt, makeRichText } from '../../src/cell';

describe('getRichTextFontAt', () => {
  it('returns the font of the covering run within a single styled run', () => {
    const rt = makeRichText([{ text: 'hello', font: { b: true, sz: 12 } }]);
    expect(getRichTextFontAt(rt, 0)).toEqual({ b: true, sz: 12 });
    expect(getRichTextFontAt(rt, 4)).toEqual({ b: true, sz: 12 });
  });

  it('returns the right font when index falls in different runs', () => {
    const rt = makeRichText([
      { text: 'AB', font: { b: true } },
      { text: 'cd', font: { i: true } },
      { text: 'EF', font: { u: 'single' } },
    ]);
    expect(getRichTextFontAt(rt, 0)).toEqual({ b: true });
    expect(getRichTextFontAt(rt, 1)).toEqual({ b: true });
    expect(getRichTextFontAt(rt, 2)).toEqual({ i: true });
    expect(getRichTextFontAt(rt, 3)).toEqual({ i: true });
    expect(getRichTextFontAt(rt, 4)).toEqual({ u: 'single' });
    expect(getRichTextFontAt(rt, 5)).toEqual({ u: 'single' });
  });

  it('returns undefined when the covering run has no font', () => {
    const rt = makeRichText([{ text: 'abc' }]);
    expect(getRichTextFontAt(rt, 0)).toBeUndefined();
    expect(getRichTextFontAt(rt, 2)).toBeUndefined();
  });

  it('returns undefined for out-of-range indices', () => {
    const rt = makeRichText([{ text: 'abc', font: { b: true } }]);
    expect(getRichTextFontAt(rt, -1)).toBeUndefined();
    expect(getRichTextFontAt(rt, 3)).toBeUndefined();
    expect(getRichTextFontAt(rt, 100)).toBeUndefined();
    expect(getRichTextFontAt(makeRichText([]), 0)).toBeUndefined();
  });
});
