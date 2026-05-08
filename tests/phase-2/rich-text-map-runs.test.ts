// Tests for the public `mapRichTextRuns(rt, fn)` writer-friendly map helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, mapRichTextRuns, richTextToString } from '../../src/cell';

describe('mapRichTextRuns', () => {
  it('applies fn to every run and returns a new frozen RichText', () => {
    const rt = makeRichText([{ text: 'a' }, { text: 'b' }, { text: 'c' }]);
    const out = mapRichTextRuns(rt, (r) => ({ text: r.text, font: { b: true } }));
    expect(out.length).toBe(3);
    expect(richTextToString(out)).toBe('abc');
    for (const r of out) expect(r.font).toEqual({ b: true });
    expect(Object.isFrozen(out)).toBe(true);
    expect(rt[0]?.font).toBeUndefined();
  });

  it('passes the run index to fn', () => {
    const rt = makeRichText([{ text: 'x' }, { text: 'y' }, { text: 'z' }]);
    const seen: number[] = [];
    const out = mapRichTextRuns(rt, (r, i) => {
      seen.push(i);
      return { text: `${r.text}${i}` };
    });
    expect(seen).toEqual([0, 1, 2]);
    expect(richTextToString(out)).toBe('x0y1z2');
  });

  it('returns an empty RichText when input is empty', () => {
    const out = mapRichTextRuns(makeRichText([]), () => ({ text: 'never' }));
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });
});
