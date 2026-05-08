// Tests for the public `richTextEqual(a, b)` predicate.

import { describe, expect, it } from 'vitest';
import { makeRichText, richTextEqual } from '../../src/cell';

describe('richTextEqual', () => {
  it('returns true for the same reference', () => {
    const rt = makeRichText([{ text: 'a', font: { b: true } }]);
    expect(richTextEqual(rt, rt)).toBe(true);
  });

  it('returns true for structurally identical RichText values built independently', () => {
    const a = makeRichText([
      { text: 'AB', font: { b: true } },
      { text: 'cd' },
    ]);
    const b = makeRichText([
      { text: 'AB', font: { b: true } },
      { text: 'cd' },
    ]);
    expect(richTextEqual(a, b)).toBe(true);
    expect(richTextEqual(makeRichText([]), makeRichText([]))).toBe(true);
  });

  it('returns false when run text differs', () => {
    const a = makeRichText([{ text: 'AB', font: { b: true } }]);
    const b = makeRichText([{ text: 'AC', font: { b: true } }]);
    expect(richTextEqual(a, b)).toBe(false);
  });

  it('returns false when font differs (presence and value)', () => {
    const noFont = makeRichText([{ text: 'AB' }]);
    const withFont = makeRichText([{ text: 'AB', font: { b: true } }]);
    expect(richTextEqual(noFont, withFont)).toBe(false);
    const otherFont = makeRichText([{ text: 'AB', font: { i: true } }]);
    expect(richTextEqual(withFont, otherFont)).toBe(false);
  });

  it('returns false when run counts differ', () => {
    const a = makeRichText([{ text: 'A' }, { text: 'B' }]);
    const b = makeRichText([{ text: 'AB' }]);
    expect(richTextEqual(a, b)).toBe(false);
  });
});
