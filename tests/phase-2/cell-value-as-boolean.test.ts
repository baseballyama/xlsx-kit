// Tests for cellValueAsBoolean.

import { describe, expect, it } from 'vitest';
import {
  cellValueAsBoolean,
  makeDurationValue,
  makeErrorValue,
} from '../../src/xlsx/cell/cell';
import { makeRichText } from '../../src/xlsx/cell/rich-text';

describe('cellValueAsBoolean', () => {
  it('booleans pass through', () => {
    expect(cellValueAsBoolean(true)).toBe(true);
    expect(cellValueAsBoolean(false)).toBe(false);
  });

  it('numeric 0 → false; any other finite number → true', () => {
    expect(cellValueAsBoolean(0)).toBe(false);
    expect(cellValueAsBoolean(1)).toBe(true);
    expect(cellValueAsBoolean(-3.14)).toBe(true);
  });

  it('NaN / Infinity → undefined', () => {
    expect(cellValueAsBoolean(Number.NaN)).toBeUndefined();
    expect(cellValueAsBoolean(Number.POSITIVE_INFINITY)).toBeUndefined();
  });

  it('"true" / "false" parse case-insensitively', () => {
    expect(cellValueAsBoolean('true')).toBe(true);
    expect(cellValueAsBoolean('TRUE')).toBe(true);
    expect(cellValueAsBoolean('False')).toBe(false);
    expect(cellValueAsBoolean('FALSE')).toBe(false);
  });

  it('arbitrary strings → undefined', () => {
    expect(cellValueAsBoolean('hello')).toBeUndefined();
    expect(cellValueAsBoolean('')).toBeUndefined();
    expect(cellValueAsBoolean('1')).toBeUndefined();
  });

  it('formula cell with cached boolean passes through', () => {
    expect(
      cellValueAsBoolean({ kind: 'formula', t: 'normal', formula: 'TRUE()', cachedValue: true }),
    ).toBe(true);
  });

  it('formula cell without cached boolean → undefined', () => {
    expect(cellValueAsBoolean({ kind: 'formula', t: 'normal', formula: 'A1', cachedValue: 7 })).toBeUndefined();
    expect(cellValueAsBoolean({ kind: 'formula', t: 'normal', formula: 'A1' })).toBeUndefined();
  });

  it('null / Date / error / duration / rich-text → undefined', () => {
    expect(cellValueAsBoolean(null)).toBeUndefined();
    expect(cellValueAsBoolean(new Date())).toBeUndefined();
    expect(cellValueAsBoolean(makeErrorValue('#REF!'))).toBeUndefined();
    expect(cellValueAsBoolean(makeDurationValue(100))).toBeUndefined();
    expect(cellValueAsBoolean({ kind: 'rich-text', runs: makeRichText([{ text: 'x' }]) })).toBeUndefined();
  });
});
