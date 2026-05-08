// Tests for the public `richTextToUpperCase(rt)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextToString, richTextToUpperCase } from '../../src/cell';

describe('richTextToUpperCase', () => {
  it('uppercases a single run while preserving its font', () => {
    const rt = makeRichText([{ text: 'Hello world', font: { i: true } }]);
    const out = richTextToUpperCase(rt);
    expect(out.length).toBe(1);
    expect(out[0]?.text).toBe('HELLO WORLD');
    expect(out[0]?.font).toEqual({ i: true });
  });

  it('uppercases each run across multiple runs', () => {
    const rt = makeRichText([
      { text: 'ab', font: { b: true } },
      { text: 'Cd', font: { i: true } },
      { text: 'ef' },
    ]);
    const out = richTextToUpperCase(rt);
    expect(richTextToString(out)).toBe('ABCDEF');
    expect(out[0]?.font).toEqual({ b: true });
    expect(out[1]?.font).toEqual({ i: true });
    expect(out[2]?.font).toBeUndefined();
  });

  it('uppercases non-ASCII characters', () => {
    const rt = makeRichText([{ text: 'äöüß' }]);
    const out = richTextToUpperCase(rt);
    expect(out[0]?.text).toBe('ÄÖÜSS');
  });

  it('returns an empty RichText when the input is empty', () => {
    const out = richTextToUpperCase(makeRichText([]));
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });
});
