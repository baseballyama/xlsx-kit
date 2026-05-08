// Tests for the public `clearFontsInRichText(rt)` helper.

import { describe, expect, it } from 'vitest';
import { clearFontsInRichText, makeRichText, richTextToString } from '../../src/cell';

describe('clearFontsInRichText', () => {
  it('removes the font from a single styled run', () => {
    const rt = makeRichText([{ text: 'hello', font: { b: true, sz: 14 } }]);
    const out = clearFontsInRichText(rt);
    expect(out.length).toBe(1);
    expect(out[0]?.text).toBe('hello');
    expect(out[0]?.font).toBeUndefined();
  });

  it('removes fonts from a mix of styled and unstyled runs', () => {
    const rt = makeRichText([
      { text: 'AB', font: { b: true } },
      { text: 'cd' },
      { text: 'EF', font: { i: true } },
    ]);
    const out = clearFontsInRichText(rt);
    expect(richTextToString(out)).toBe('ABcdEF');
    for (const r of out) expect(r.font).toBeUndefined();
  });

  it('returns an empty RichText when the input is empty', () => {
    const out = clearFontsInRichText(makeRichText([]));
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });
});
