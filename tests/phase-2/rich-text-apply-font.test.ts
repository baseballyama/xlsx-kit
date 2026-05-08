// Tests for the public `applyFontToRichText(rt, font)` bulk-merge helper.

import { describe, expect, it } from 'vitest';
import { applyFontToRichText, makeRichText } from '../../src/cell';

describe('applyFontToRichText', () => {
  it('applies a common font to every run that lacks one', () => {
    const rt = makeRichText([{ text: 'a' }, { text: 'b' }]);
    const out = applyFontToRichText(rt, { b: true, sz: 12 });
    expect(out.length).toBe(2);
    expect(out[0]?.font).toEqual({ b: true, sz: 12 });
    expect(out[1]?.font).toEqual({ b: true, sz: 12 });
    expect(Object.isFrozen(out)).toBe(true);
  });

  it('per-run font overrides the common font on overlapping fields', () => {
    const rt = makeRichText([
      { text: 'common' },
      { text: 'override', font: { b: false, i: true } },
    ]);
    const out = applyFontToRichText(rt, { b: true, sz: 14 });
    expect(out[0]?.font).toEqual({ b: true, sz: 14 });
    expect(out[1]?.font).toEqual({ b: false, sz: 14, i: true });
  });

  it('returns an empty RichText when input is empty', () => {
    const out = applyFontToRichText(makeRichText([]), { b: true });
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });
});
