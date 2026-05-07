// Tests for shiftRangeStr — A1-string range translation.

import { describe, expect, it } from 'vitest';
import { shiftRangeStr } from '../../src/worksheet/cell-range';

describe('shiftRangeStr', () => {
  it('shifts both row and column with positive offsets', () => {
    expect(shiftRangeStr('A1:C5', 1, 1)).toBe('B2:D6');
  });

  it('shifts with negative offsets (move up / left)', () => {
    expect(shiftRangeStr('B2:D6', -1, -1)).toBe('A1:C5');
  });

  it('handles single-cell refs (still serialised as A1:A1 form? rangeToString choice)', () => {
    // shiftRange returns a CellRange and rangeToString collapses single-cell.
    expect(shiftRangeStr('A1', 4, 0)).toBe('A5');
  });

  it('returns the same range string for (0, 0)', () => {
    expect(shiftRangeStr('A1:C5', 0, 0)).toBe('A1:C5');
  });

  it('throws when the resulting range falls outside the OOXML grid', () => {
    // Excel max row = 1048576; shifting A1 by 1048576 lands at row 1048577 (invalid).
    expect(() => shiftRangeStr('A1', 1048576, 0)).toThrow();
    expect(() => shiftRangeStr('A1', -1, 0)).toThrow();
  });
});
