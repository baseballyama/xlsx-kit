// Tests for the public `trimStartRichText(rt)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextToString, trimStartRichText } from '../../src/cell';

describe('trimStartRichText', () => {
  it('strips only leading whitespace within a single run, preserving font', () => {
    const rt = makeRichText([{ text: '  hello\t', font: { b: true } }]);
    const out = trimStartRichText(rt);
    expect(richTextToString(out)).toBe('hello\t');
    expect(out.length).toBe(1);
    expect(out[0]?.font).toEqual({ b: true });
  });

  it('strips leading whitespace that spans multiple runs', () => {
    const rt = makeRichText([
      { text: '  \t', font: { b: true } },
      { text: 'hello', font: { i: true } },
      { text: ' world  ' },
    ]);
    const out = trimStartRichText(rt);
    expect(richTextToString(out)).toBe('hello world  ');
    expect(out[0]?.font).toEqual({ i: true });
  });

  it('keeps trailing whitespace intact', () => {
    const rt = makeRichText([{ text: ' a\t' }]);
    const out = trimStartRichText(rt);
    expect(richTextToString(out)).toBe('a\t');
  });

  it('returns an empty RichText when every character is whitespace', () => {
    const rt = makeRichText([{ text: ' \t' }, { text: '\r\n  ' }]);
    const out = trimStartRichText(rt);
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });
});
