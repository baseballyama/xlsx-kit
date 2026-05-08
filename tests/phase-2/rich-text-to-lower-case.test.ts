// Tests for the public `richTextToLowerCase(rt)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextToLowerCase, richTextToString } from '../../src/cell';

describe('richTextToLowerCase', () => {
  it('lowercases a single run while preserving its font', () => {
    const rt = makeRichText([{ text: 'Hello WORLD', font: { b: true } }]);
    const out = richTextToLowerCase(rt);
    expect(out.length).toBe(1);
    expect(out[0]?.text).toBe('hello world');
    expect(out[0]?.font).toEqual({ b: true });
  });

  it('lowercases each run across multiple runs', () => {
    const rt = makeRichText([
      { text: 'AB', font: { b: true } },
      { text: 'cD', font: { i: true } },
      { text: 'EF' },
    ]);
    const out = richTextToLowerCase(rt);
    expect(richTextToString(out)).toBe('abcdef');
    expect(out[0]?.font).toEqual({ b: true });
    expect(out[1]?.font).toEqual({ i: true });
    expect(out[2]?.font).toBeUndefined();
  });

  it('lowercases non-ASCII characters', () => {
    const rt = makeRichText([{ text: 'ÄÖÜß' }]);
    const out = richTextToLowerCase(rt);
    expect(out[0]?.text).toBe('äöüß');
  });

  it('returns an empty RichText when the input is empty', () => {
    const out = richTextToLowerCase(makeRichText([]));
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });
});
