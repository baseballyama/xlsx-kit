// Tests for the public `setFontOnRichText(rt, font)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextToString, setFontOnRichText } from '../../src/cell';

describe('setFontOnRichText', () => {
  it('overrides existing per-run fonts', () => {
    const rt = makeRichText([
      { text: 'AB', font: { b: true } },
      { text: 'cd', font: { i: true } },
    ]);
    const out = setFontOnRichText(rt, { u: 'single' });
    expect(richTextToString(out)).toBe('ABcd');
    for (const r of out) expect(r.font).toEqual({ u: 'single' });
  });

  it('applies the font to runs that previously had none', () => {
    const rt = makeRichText([{ text: 'AB' }, { text: 'cd' }]);
    const out = setFontOnRichText(rt, { b: true });
    for (const r of out) expect(r.font).toEqual({ b: true });
  });

  it('returns an empty RichText when the input is empty', () => {
    const out = setFontOnRichText(makeRichText([]), { b: true });
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });
});
