// Tests for the public `trimEndRichText(rt)` helper.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextToString, trimEndRichText } from '../../src/cell';

describe('trimEndRichText', () => {
  it('strips only trailing whitespace within a single run, preserving font', () => {
    const rt = makeRichText([{ text: '\thello  ', font: { b: true } }]);
    const out = trimEndRichText(rt);
    expect(richTextToString(out)).toBe('\thello');
    expect(out.length).toBe(1);
    expect(out[0]?.font).toEqual({ b: true });
  });

  it('strips trailing whitespace that spans multiple runs', () => {
    const rt = makeRichText([
      { text: '  hello', font: { b: true } },
      { text: ' world', font: { i: true } },
      { text: ' \t  \r\n' },
    ]);
    const out = trimEndRichText(rt);
    expect(richTextToString(out)).toBe('  hello world');
    expect(out[out.length - 1]?.font).toEqual({ i: true });
  });

  it('keeps leading whitespace intact', () => {
    const rt = makeRichText([{ text: '\ta ' }]);
    const out = trimEndRichText(rt);
    expect(richTextToString(out)).toBe('\ta');
  });

  it('returns an empty RichText when every character is whitespace', () => {
    const rt = makeRichText([{ text: ' \t' }, { text: '\r\n  ' }]);
    const out = trimEndRichText(rt);
    expect(out.length).toBe(0);
    expect(Object.isFrozen(out)).toBe(true);
  });
});
