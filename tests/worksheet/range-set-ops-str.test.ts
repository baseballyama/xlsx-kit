// Tests for unionRangeStr / intersectionRangeStr — A1 set operations.

import { describe, expect, it } from 'vitest';
import { intersectionRangeStr, unionRangeStr } from '../../src/worksheet/cell-range';

describe('unionRangeStr', () => {
  it('returns the bounding box of two overlapping ranges', () => {
    expect(unionRangeStr('A1:C5', 'B3:D7')).toBe('A1:D7');
  });

  it('returns the bounding box even for disjoint ranges (always non-null)', () => {
    expect(unionRangeStr('A1:B2', 'D4:E5')).toBe('A1:E5');
  });

  it('returns the same range for two identical inputs', () => {
    expect(unionRangeStr('A1:C5', 'A1:C5')).toBe('A1:C5');
  });

  it('handles single-cell refs on either side', () => {
    expect(unionRangeStr('A1', 'C5')).toBe('A1:C5');
  });
});

describe('intersectionRangeStr', () => {
  it('returns the shared sub-range when the inputs overlap', () => {
    expect(intersectionRangeStr('A1:C5', 'B3:D7')).toBe('B3:C5');
  });

  it('returns undefined when the inputs are disjoint', () => {
    expect(intersectionRangeStr('A1:B2', 'D4:E5')).toBeUndefined();
  });

  it('returns the smaller range when one fully contains the other', () => {
    expect(intersectionRangeStr('A1:E10', 'B2:C3')).toBe('B2:C3');
  });

  it('returns the same range for two identical inputs', () => {
    expect(intersectionRangeStr('A1:C5', 'A1:C5')).toBe('A1:C5');
  });

  it('returns the touching cell for ranges sharing a single boundary corner', () => {
    expect(intersectionRangeStr('A1:B2', 'B2:C3')).toBe('B2');
  });
});
