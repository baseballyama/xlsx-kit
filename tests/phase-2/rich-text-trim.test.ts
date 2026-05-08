// Tests for the public `trimRichText(rt)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextToString, trimRichText } from '../../src/cell';

describe('trimRichText', () => {
  it('trims leading and trailing whitespace within a single run, preserving font', () => {
    const rt = makeRichText([{ text: '  hello\t', font: { b: true } }]);
    const out = trimRichText(rt);
    expect(richTextToString(out)).toBe('hello');
    expect(out.length).toBe(1);
    expect(out[0]?.font).toEqual({ b: true });
  });

  it('trims whitespace that spans multiple runs', () => {
    const rt = makeRichText([
      { text: '  \t', font: { b: true } },
      { text: 'hello', font: { i: true } },
      { text: ' world  ', font: { u: 'single' } },
      { text: '\r\n' },
    ]);
    const out = trimRichText(rt);
    expect(richTextToString(out)).toBe('hello world');
    expect(out[0]?.font).toEqual({ i: true });
    expect(out[out.length - 1]?.font).toEqual({ u: 'single' });
  });

  it('returns an empty RichText when every character is ASCII whitespace', () => {
    const rt = makeRichText([{ text: '  \t' }, { text: '\r\n  ' }]);
    const out = trimRichText(rt);
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });

  it('keeps internal whitespace untouched', () => {
    const rt = makeRichText([{ text: '  a  b  c  ' }]);
    const out = trimRichText(rt);
    expect(richTextToString(out)).toBe('a  b  c');
  });
});
