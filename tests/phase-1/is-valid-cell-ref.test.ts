// Tests for isValidCellRef predicate.

import { describe, expect, it } from 'vitest';
import { isValidCellRef } from '../../src/utils/coordinate';

describe('isValidCellRef', () => {
  it('accepts plain A1-style refs', () => {
    expect(isValidCellRef('A1')).toBe(true);
    expect(isValidCellRef('Z99')).toBe(true);
    expect(isValidCellRef('XFD1048576')).toBe(true);
    expect(isValidCellRef('a1')).toBe(true); // case-insensitive
  });

  it('rejects refs with $ absolute markers', () => {
    expect(isValidCellRef('$A$1')).toBe(false);
    expect(isValidCellRef('A$1')).toBe(false);
    expect(isValidCellRef('$A1')).toBe(false);
  });

  it('rejects ranges', () => {
    expect(isValidCellRef('A1:B2')).toBe(false);
    expect(isValidCellRef('A:A')).toBe(false);
    expect(isValidCellRef('1:1')).toBe(false);
  });

  it('rejects whitespace and empty strings', () => {
    expect(isValidCellRef('')).toBe(false);
    expect(isValidCellRef(' A1')).toBe(false);
    expect(isValidCellRef('A1 ')).toBe(false);
  });

  it('rejects out-of-bounds rows / cols', () => {
    expect(isValidCellRef('A0')).toBe(false);
    expect(isValidCellRef('A1048577')).toBe(false); // row max + 1
    expect(isValidCellRef('XFE1')).toBe(false); // col max + 1
  });

  it('rejects non-string input (also typed-narrows in TS)', () => {
    expect(isValidCellRef(42)).toBe(false);
    expect(isValidCellRef(null)).toBe(false);
    expect(isValidCellRef(undefined)).toBe(false);
  });
});
