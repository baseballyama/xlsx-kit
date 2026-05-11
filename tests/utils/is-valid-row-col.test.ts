// Tests for isValidColumnLetter / isValidRowNumber / isValidColumnNumber.

import { describe, expect, it } from 'vitest';
import {
  isValidColumnLetter,
  isValidColumnNumber,
  isValidRowNumber,
} from '../../src/utils/coordinate';

describe('isValidColumnLetter', () => {
  it('accepts A..XFD (case-insensitive)', () => {
    expect(isValidColumnLetter('A')).toBe(true);
    expect(isValidColumnLetter('Z')).toBe(true);
    expect(isValidColumnLetter('AA')).toBe(true);
    expect(isValidColumnLetter('XFD')).toBe(true);
    expect(isValidColumnLetter('xfd')).toBe(true);
  });

  it('rejects empty / over-length / non-letters', () => {
    expect(isValidColumnLetter('')).toBe(false);
    expect(isValidColumnLetter('XFDA')).toBe(false);
    expect(isValidColumnLetter('A1')).toBe(false);
    expect(isValidColumnLetter(' A')).toBe(false);
  });

  it('rejects out-of-bound XFE', () => {
    expect(isValidColumnLetter('XFE')).toBe(false);
  });

  it('rejects non-string', () => {
    expect(isValidColumnLetter(1)).toBe(false);
    expect(isValidColumnLetter(null)).toBe(false);
  });
});

describe('isValidRowNumber', () => {
  it('accepts integers in [1, 1048576]', () => {
    expect(isValidRowNumber(1)).toBe(true);
    expect(isValidRowNumber(1048576)).toBe(true);
  });

  it('rejects 0 and negatives', () => {
    expect(isValidRowNumber(0)).toBe(false);
    expect(isValidRowNumber(-1)).toBe(false);
  });

  it('rejects > 1048576', () => {
    expect(isValidRowNumber(1048577)).toBe(false);
  });

  it('rejects non-integers / NaN / Infinity', () => {
    expect(isValidRowNumber(1.5)).toBe(false);
    expect(isValidRowNumber(Number.NaN)).toBe(false);
    expect(isValidRowNumber(Number.POSITIVE_INFINITY)).toBe(false);
  });

  it('rejects non-number', () => {
    expect(isValidRowNumber('1')).toBe(false);
    expect(isValidRowNumber(null)).toBe(false);
  });
});

describe('isValidColumnNumber', () => {
  it('accepts integers in [1, 16384]', () => {
    expect(isValidColumnNumber(1)).toBe(true);
    expect(isValidColumnNumber(16384)).toBe(true);
  });

  it('rejects 0 and > 16384', () => {
    expect(isValidColumnNumber(0)).toBe(false);
    expect(isValidColumnNumber(16385)).toBe(false);
  });

  it('rejects non-integer', () => {
    expect(isValidColumnNumber(2.5)).toBe(false);
  });
});
