// Tests for expandRangeStr — A1 range bottom/right resize.

import { describe, expect, it } from 'vitest';
import { expandRangeStr } from '../../src/xlsx/worksheet/cell-range';

describe('expandRangeStr', () => {
  it('extends the bottom-right corner by the deltas', () => {
    expect(expandRangeStr('A1:C5', 2, 1)).toBe('A1:D7');
  });

  it('returns the same range for (0, 0)', () => {
    expect(expandRangeStr('A1:C5', 0, 0)).toBe('A1:C5');
  });

  it('extends only one axis when the other delta is 0', () => {
    expect(expandRangeStr('A1:C5', 0, 2)).toBe('A1:E5');
    expect(expandRangeStr('A1:C5', 3, 0)).toBe('A1:C8');
  });

  it('promotes a single-cell ref to a multi-cell range', () => {
    expect(expandRangeStr('A1', 4, 2)).toBe('A1:C5');
  });

  it('shrinks the range when deltas are negative', () => {
    expect(expandRangeStr('A1:C5', -2, -1)).toBe('A1:B3');
  });

  it('throws when the result would have zero or negative dimensions', () => {
    expect(() => expandRangeStr('A1:C5', -5, 0)).toThrow();
    expect(() => expandRangeStr('A1:C5', 0, -3)).toThrow();
  });

  it('throws on non-integer deltas', () => {
    expect(() => expandRangeStr('A1:C5', 1.5, 0)).toThrow(/integers/);
  });
});
