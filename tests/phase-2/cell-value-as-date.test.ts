// Tests for cellValueAsDate.

import { describe, expect, it } from 'vitest';
import {
  cellValueAsDate,
  makeDurationValue,
  makeErrorValue,
} from '../../src/cell/cell';
import { makeRichText } from '../../src/cell/rich-text';

describe('cellValueAsDate', () => {
  it('Date passes through', () => {
    const d = new Date('2024-01-15T00:00:00Z');
    expect(cellValueAsDate(d)).toBe(d);
  });

  it('ISO string parses', () => {
    const out = cellValueAsDate('2024-01-15T00:00:00Z');
    expect(out?.toISOString()).toBe('2024-01-15T00:00:00.000Z');
  });

  it('non-date string returns undefined', () => {
    expect(cellValueAsDate('hello')).toBeUndefined();
  });

  it('empty string returns undefined', () => {
    expect(cellValueAsDate('')).toBeUndefined();
  });

  it('duration converts via new Date(ms)', () => {
    const out = cellValueAsDate(makeDurationValue(60000));
    expect(out?.getTime()).toBe(60000);
  });

  it('numbers are NOT auto-treated as Excel serials', () => {
    expect(cellValueAsDate(45000)).toBeUndefined();
    expect(cellValueAsDate(0)).toBeUndefined();
  });

  it('null / boolean / error / rich-text → undefined', () => {
    expect(cellValueAsDate(null)).toBeUndefined();
    expect(cellValueAsDate(true)).toBeUndefined();
    expect(cellValueAsDate(makeErrorValue('#REF!'))).toBeUndefined();
    expect(cellValueAsDate({ kind: 'rich-text', runs: makeRichText([{ text: 'hi' }]) })).toBeUndefined();
  });
});
