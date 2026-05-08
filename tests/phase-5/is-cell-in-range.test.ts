// Tests for isCellInRange — A1-string convenience predicate.

import { describe, expect, it } from 'vitest';
import { isCellInRange } from '../../src/xlsx/worksheet/cell-range';

describe('isCellInRange', () => {
  it('returns true for cells inside the range', () => {
    expect(isCellInRange('B3', 'A1:C5')).toBe(true);
    expect(isCellInRange('A1', 'A1:A1')).toBe(true);
  });

  it('treats range boundaries as inclusive', () => {
    expect(isCellInRange('A1', 'A1:C5')).toBe(true);
    expect(isCellInRange('C5', 'A1:C5')).toBe(true);
    expect(isCellInRange('A5', 'A1:C5')).toBe(true);
    expect(isCellInRange('C1', 'A1:C5')).toBe(true);
  });

  it('returns false for cells outside the range', () => {
    expect(isCellInRange('A6', 'A1:C5')).toBe(false);
    expect(isCellInRange('D1', 'A1:C5')).toBe(false);
    expect(isCellInRange('Z99', 'A1:C5')).toBe(false);
  });

  it('throws on a malformed cell ref', () => {
    expect(() => isCellInRange('not-a-cell', 'A1:C5')).toThrow();
  });

  it('throws on a malformed range ref', () => {
    expect(() => isCellInRange('A1', 'not-a-range')).toThrow();
  });
});
