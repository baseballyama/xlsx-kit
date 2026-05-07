// Tests for isValidRangeRef predicate.

import { describe, expect, it } from 'vitest';
import { isValidRangeRef } from '../../src/utils/coordinate';

describe('isValidRangeRef', () => {
  it('accepts single cells', () => {
    expect(isValidRangeRef('A1')).toBe(true);
    expect(isValidRangeRef('XFD1048576')).toBe(true);
  });

  it('accepts two-corner ranges', () => {
    expect(isValidRangeRef('A1:B5')).toBe(true);
    expect(isValidRangeRef('A1:A1')).toBe(true);
  });

  it('accepts whole-column ranges', () => {
    expect(isValidRangeRef('A:A')).toBe(true);
    expect(isValidRangeRef('A:Z')).toBe(true);
    expect(isValidRangeRef('XFD:XFD')).toBe(true);
  });

  it('accepts whole-row ranges', () => {
    expect(isValidRangeRef('1:1')).toBe(true);
    expect(isValidRangeRef('1:1048576')).toBe(true);
  });

  it('rejects $ markers, whitespace, and empty', () => {
    expect(isValidRangeRef('$A$1:$B$2')).toBe(false);
    expect(isValidRangeRef('A1 :B2')).toBe(false);
    expect(isValidRangeRef('')).toBe(false);
  });

  it('rejects out-of-bound rows / cols', () => {
    expect(isValidRangeRef('A0:B0')).toBe(false);
    expect(isValidRangeRef('A1:A1048577')).toBe(false);
    expect(isValidRangeRef('XFE:XFE')).toBe(false);
  });

  it('rejects non-string', () => {
    expect(isValidRangeRef(42)).toBe(false);
    expect(isValidRangeRef(null)).toBe(false);
  });

  it('rejects multi-range / sheet-prefix forms', () => {
    expect(isValidRangeRef('A1:B2 C3:D4')).toBe(false);
    expect(isValidRangeRef("Sheet1!A1")).toBe(false);
  });
});
