// Tests for rangeDimensionsStr — A1 range → { rows, cols }.

import { describe, expect, it } from 'vitest';
import { rangeDimensionsStr } from '../../src/xlsx/worksheet/cell-range';

describe('rangeDimensionsStr', () => {
  it('returns {rows:1, cols:1} for a single-cell ref', () => {
    expect(rangeDimensionsStr('A1')).toEqual({ rows: 1, cols: 1 });
  });

  it('returns the row + col span for a rectangular range', () => {
    expect(rangeDimensionsStr('A1:C5')).toEqual({ rows: 5, cols: 3 });
    expect(rangeDimensionsStr('B2:D4')).toEqual({ rows: 3, cols: 3 });
  });

  it('returns {rows:N, cols:1} for a single-column range', () => {
    expect(rangeDimensionsStr('A1:A10')).toEqual({ rows: 10, cols: 1 });
  });

  it('returns {rows:1, cols:N} for a single-row range', () => {
    expect(rangeDimensionsStr('A1:E1')).toEqual({ rows: 1, cols: 5 });
  });

  it('throws on malformed input', () => {
    expect(() => rangeDimensionsStr('not-a-range')).toThrow();
  });
});
