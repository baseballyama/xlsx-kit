// Tests for the public `isEmptyRichText(rt)` predicate.

import { describe, expect, it } from 'vitest';
import { isEmptyRichText, makeRichText } from '../../src/cell';

describe('isEmptyRichText', () => {
  it('returns true for a RichText with no runs', () => {
    expect(isEmptyRichText(makeRichText([]))).toBe(true);
  });

  it('returns true when every run carries an empty text', () => {
    expect(isEmptyRichText(makeRichText([{ text: '' }]))).toBe(true);
    expect(
      isEmptyRichText(
        makeRichText([
          { text: '', font: { b: true } },
          { text: '' },
        ]),
      ),
    ).toBe(true);
  });

  it('returns false when any run carries non-empty text', () => {
    expect(isEmptyRichText(makeRichText([{ text: 'a' }]))).toBe(false);
    expect(
      isEmptyRichText(
        makeRichText([
          { text: '', font: { b: true } },
          { text: 'x' },
        ]),
      ),
    ).toBe(false);
  });
});
