// Tests for the public `mergeAdjacentRichTextRuns(rt)` compaction helper.

import { describe, expect, it } from 'vitest';
import {
  makeRichText,
  mergeAdjacentRichTextRuns,
  richTextToString,
  splitRichTextRuns,
} from '../../src/cell';

describe('mergeAdjacentRichTextRuns', () => {
  it('merges adjacent font-less runs into one', () => {
    const rt = makeRichText([{ text: 'a' }, { text: 'b' }, { text: 'c' }]);
    const out = mergeAdjacentRichTextRuns(rt);
    expect(out.length).toBe(1);
    expect(out[0]?.text).toBe('abc');
    expect(out[0]?.font).toBeUndefined();
  });

  it('merges adjacent runs that share an identical font', () => {
    const rt = makeRichText([
      { text: 'a', font: { b: true, sz: 12 } },
      { text: 'b', font: { b: true, sz: 12 } },
      { text: 'c', font: { i: true } },
    ]);
    const out = mergeAdjacentRichTextRuns(rt);
    expect(out.length).toBe(2);
    expect(out[0]?.text).toBe('ab');
    expect(out[0]?.font).toEqual({ b: true, sz: 12 });
    expect(out[1]?.text).toBe('c');
    expect(out[1]?.font).toEqual({ i: true });
  });

  it('keeps adjacent runs with different fonts separate', () => {
    const rt = makeRichText([
      { text: 'a', font: { b: true } },
      { text: 'b', font: { i: true } },
    ]);
    const out = mergeAdjacentRichTextRuns(rt);
    expect(out.length).toBe(2);
    expect(out[0]?.font).toEqual({ b: true });
    expect(out[1]?.font).toEqual({ i: true });
  });

  it('split → merge round-trip compacts same-font runs back together', () => {
    const rt = makeRichText([
      { text: 'hello', font: { b: true } },
      { text: 'world', font: { b: true } },
    ]);
    const split = splitRichTextRuns(rt);
    expect(split.length).toBe(10);
    const merged = mergeAdjacentRichTextRuns(split);
    expect(merged.length).toBe(1);
    expect(richTextToString(merged)).toBe('helloworld');
    expect(merged[0]?.font).toEqual({ b: true });
  });
});
