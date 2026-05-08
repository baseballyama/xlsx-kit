// Tests for rangesOverlapStr — A1-string range-overlap predicate.

import { describe, expect, it } from 'vitest';
import { rangesOverlapStr } from '../../src/xlsx/worksheet/cell-range';

describe('rangesOverlapStr', () => {
  it('returns true for ranges that share at least one cell', () => {
    expect(rangesOverlapStr('A1:C5', 'B3:D7')).toBe(true);
  });

  it('returns true for fully-contained ranges (containment ⊆ overlap)', () => {
    expect(rangesOverlapStr('A1:E10', 'B2:C3')).toBe(true);
  });

  it('returns false for disjoint ranges separated by a gap', () => {
    expect(rangesOverlapStr('A1:B2', 'D4:E5')).toBe(false);
  });

  it('returns true for touching boundaries (boundary-inclusive)', () => {
    // A1:B2 and B2:C3 share B2 → overlap.
    expect(rangesOverlapStr('A1:B2', 'B2:C3')).toBe(true);
  });

  it('accepts single-cell refs on either side', () => {
    expect(rangesOverlapStr('B3', 'A1:C5')).toBe(true);
    expect(rangesOverlapStr('Z9', 'A1:B2')).toBe(false);
  });

  it('throws on malformed input', () => {
    expect(() => rangesOverlapStr('not-a-range', 'A1:B2')).toThrow();
    expect(() => rangesOverlapStr('A1:B2', 'not-a-range')).toThrow();
  });
});
