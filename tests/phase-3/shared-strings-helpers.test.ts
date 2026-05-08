// Tests for getSharedStringIndex / getSharedStringAt / sharedStringCount.

import { describe, expect, it } from 'vitest';
import {
  addSharedString,
  getSharedStringAt,
  getSharedStringIndex,
  makeSharedStrings,
  sharedStringCount,
} from '../../src/xlsx/workbook/shared-strings';

describe('SST helpers', () => {
  it('count grows with each unique add', () => {
    const sst = makeSharedStrings();
    expect(sharedStringCount(sst)).toBe(0);
    addSharedString(sst, 'foo');
    addSharedString(sst, 'bar');
    expect(sharedStringCount(sst)).toBe(2);
  });

  it('repeated adds dedupe so count stays put', () => {
    const sst = makeSharedStrings();
    addSharedString(sst, 'foo');
    addSharedString(sst, 'foo');
    addSharedString(sst, 'foo');
    expect(sharedStringCount(sst)).toBe(1);
  });

  it('getSharedStringIndex returns the index for known values', () => {
    const sst = makeSharedStrings();
    addSharedString(sst, 'a');
    const idB = addSharedString(sst, 'b');
    expect(getSharedStringIndex(sst, 'b')).toBe(idB);
    expect(getSharedStringIndex(sst, 'a')).toBe(0);
  });

  it('getSharedStringIndex returns undefined for unknown values', () => {
    const sst = makeSharedStrings();
    addSharedString(sst, 'a');
    expect(getSharedStringIndex(sst, 'missing')).toBeUndefined();
  });

  it('getSharedStringAt returns the string by index', () => {
    const sst = makeSharedStrings();
    addSharedString(sst, 'first');
    addSharedString(sst, 'second');
    expect(getSharedStringAt(sst, 0)).toBe('first');
    expect(getSharedStringAt(sst, 1)).toBe('second');
  });

  it('getSharedStringAt returns undefined for out-of-range', () => {
    const sst = makeSharedStrings();
    addSharedString(sst, 'only');
    expect(getSharedStringAt(sst, 1)).toBeUndefined();
    expect(getSharedStringAt(sst, -1)).toBeUndefined();
  });
});
