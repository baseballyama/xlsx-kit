// Tests for the public `replaceAllRichText(rt, search, replacement)` helper.

import { describe, expect, it } from 'vitest';
import {
  makeRichText,
  replaceAllRichText,
  richText,
  richTextToString,
} from '../../src/cell';

describe('replaceAllRichText', () => {
  it('replaces every occurrence within a single run', () => {
    const rt = makeRichText([{ text: 'foo bar foo bar' }]);
    const out = replaceAllRichText(rt, 'foo', 'baz');
    expect(richTextToString(out)).toBe('baz bar baz bar');
  });

  it('replaces occurrences that span multiple runs', () => {
    const rt = makeRichText([
      { text: 'aXX' },
      { text: 'XXb' },
      { text: 'cXXXXd' },
    ]);
    const out = replaceAllRichText(rt, 'XX', 'YY');
    expect(richTextToString(out)).toBe('aYYYYbcYYYYd');
  });

  it('preserves replacement run fonts on each occurrence', () => {
    const rt = makeRichText([{ text: 'a__b__c' }]);
    const replacement = richText('-', { b: true });
    const out = replaceAllRichText(rt, '__', replacement);
    expect(richTextToString(out)).toBe('a-b-c');
    const boldRuns = out.filter((r) => r.font?.b === true);
    expect(boldRuns.length).toBe(2);
  });

  it('returns the input unchanged when search is empty', () => {
    const rt = makeRichText([{ text: 'unchanged' }]);
    const out = replaceAllRichText(rt, '', 'X');
    expect(out).toBe(rt);
  });
});
