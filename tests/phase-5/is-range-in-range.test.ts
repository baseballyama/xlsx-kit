// Tests for isRangeInRange — A1-string range containment predicate.

import { describe, expect, it } from 'vitest';
import { isRangeInRange } from '../../src/xlsx/worksheet/cell-range';

describe('isRangeInRange', () => {
  it('returns true for a fully contained inner range', () => {
    expect(isRangeInRange('B2:C3', 'A1:D4')).toBe(true);
  });

  it('returns true when inner equals outer (boundary-inclusive)', () => {
    expect(isRangeInRange('A1:C5', 'A1:C5')).toBe(true);
  });

  it('returns false when the inner range only partially overlaps', () => {
    expect(isRangeInRange('A1:E5', 'C3:G7')).toBe(false);
  });

  it('returns false for a fully disjoint inner range', () => {
    expect(isRangeInRange('Z90:AA99', 'A1:B2')).toBe(false);
  });

  it('accepts single-cell refs on either side', () => {
    expect(isRangeInRange('B3', 'A1:C5')).toBe(true);
    expect(isRangeInRange('A1', 'A1')).toBe(true);
    expect(isRangeInRange('Z9', 'A1')).toBe(false);
  });

  it('throws on malformed input on either side', () => {
    expect(() => isRangeInRange('not-a-range', 'A1:C5')).toThrow();
    expect(() => isRangeInRange('A1:C5', 'not-a-range')).toThrow();
  });
});
