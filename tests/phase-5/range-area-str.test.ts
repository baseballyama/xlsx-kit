// Tests for rangeAreaStr — A1-string range area helper.

import { describe, expect, it } from 'vitest';
import { rangeAreaStr } from '../../src/worksheet/cell-range';

describe('rangeAreaStr', () => {
  it('returns 1 for a single-cell ref', () => {
    expect(rangeAreaStr('A1')).toBe(1);
    expect(rangeAreaStr('Z99')).toBe(1);
  });

  it('returns the correct count for a rectangular range', () => {
    expect(rangeAreaStr('A1:C5')).toBe(15); // 3 cols × 5 rows
    expect(rangeAreaStr('B2:D4')).toBe(9); // 3 cols × 3 rows
  });

  it('returns the row count for a single-column range', () => {
    expect(rangeAreaStr('A1:A10')).toBe(10);
  });

  it('returns the column count for a single-row range', () => {
    expect(rangeAreaStr('A1:E1')).toBe(5);
  });

  it('throws on malformed input', () => {
    expect(() => rangeAreaStr('not-a-range')).toThrow();
  });
});
