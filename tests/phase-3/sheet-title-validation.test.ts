// Tests for validateSheetTitle / isValidSheetTitle.

import { describe, expect, it } from 'vitest';
import { isValidSheetTitle, validateSheetTitle } from '../../src/workbook/workbook';

describe('validateSheetTitle', () => {
  it('accepts a normal title', () => {
    expect(validateSheetTitle('Sheet1')).toBeUndefined();
    expect(validateSheetTitle('Q3 2025')).toBeUndefined();
    expect(validateSheetTitle('日本語シート')).toBeUndefined();
  });

  it('rejects non-string / empty / too-long', () => {
    expect(validateSheetTitle(123)).toMatch(/string/);
    expect(validateSheetTitle('')).toMatch(/1\.\.31/);
    expect(validateSheetTitle('x'.repeat(32))).toMatch(/1\.\.31/);
  });

  it('rejects forbidden characters', () => {
    expect(validateSheetTitle('Sheet:1')).toMatch(/must not contain/);
    expect(validateSheetTitle('Sheet\\1')).toMatch(/must not contain/);
    expect(validateSheetTitle('Sheet/1')).toMatch(/must not contain/);
    expect(validateSheetTitle('Sheet?')).toMatch(/must not contain/);
    expect(validateSheetTitle('Sheet*')).toMatch(/must not contain/);
    expect(validateSheetTitle('Sheet[1]')).toMatch(/must not contain/);
  });

  it('rejects leading or trailing apostrophe', () => {
    expect(validateSheetTitle("'Sheet")).toMatch(/apostrophe/);
    expect(validateSheetTitle("Sheet'")).toMatch(/apostrophe/);
    expect(validateSheetTitle("Mid'dle")).toBeUndefined();
  });

  it('rejects the reserved name "History" case-insensitively', () => {
    expect(validateSheetTitle('History')).toMatch(/reserved/);
    expect(validateSheetTitle('history')).toMatch(/reserved/);
    expect(validateSheetTitle('HISTORY')).toMatch(/reserved/);
    expect(validateSheetTitle('Historical')).toBeUndefined();
  });

  it('exactly 31 characters is the inclusive upper bound', () => {
    expect(validateSheetTitle('x'.repeat(31))).toBeUndefined();
  });
});

describe('isValidSheetTitle', () => {
  it('returns true for a valid title (and narrows to string)', () => {
    expect(isValidSheetTitle('Sheet1')).toBe(true);
  });

  it('returns false for invalid titles', () => {
    expect(isValidSheetTitle('')).toBe(false);
    expect(isValidSheetTitle('Sheet:1')).toBe(false);
    expect(isValidSheetTitle(42)).toBe(false);
  });
});
