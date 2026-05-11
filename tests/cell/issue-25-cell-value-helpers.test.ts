// Tests for the #25 additions: cellValueAsString opts (dateFormat / emptyText)
// + the new cellValueAsPrimitive helper.

import { describe, expect, it } from 'vitest';
import {
  cellValueAsPrimitive,
  cellValueAsString,
  makeDurationValue,
  makeErrorValue,
} from '../../src/cell/cell';
import { makeRichText } from '../../src/cell/rich-text';

describe('cellValueAsString — options (#25)', () => {
  it('emptyText overrides the null placeholder', () => {
    expect(cellValueAsString(null, { emptyText: '—' })).toBe('—');
  });

  it('emptyText only applies to null, not to formula-with-no-cached-value', () => {
    const f = { kind: 'formula' as const, t: 'normal' as const, formula: 'A1+1' };
    // Uncached formula still returns '' regardless of emptyText.
    expect(cellValueAsString(f, { emptyText: '—' })).toBe('');
  });

  it('dateFormat overrides the Date renderer', () => {
    const d = new Date('2024-01-02T03:04:05Z');
    expect(cellValueAsString(d, { dateFormat: (v) => v.getUTCFullYear().toString() })).toBe('2024');
  });

  it('default behavior is byte-for-byte unchanged when opts is undefined', () => {
    expect(cellValueAsString(null)).toBe('');
    expect(cellValueAsString(42)).toBe('42');
    expect(cellValueAsString(new Date('2024-01-02T03:04:05Z'))).toBe('2024-01-02T03:04:05.000Z');
    expect(cellValueAsString(makeDurationValue(1500))).toBe('1500 ms');
  });
});

describe('cellValueAsPrimitive (#25)', () => {
  it('null → null', () => {
    expect(cellValueAsPrimitive(null)).toBe(null);
  });

  it('passes through string / number / boolean / Date', () => {
    expect(cellValueAsPrimitive('hi')).toBe('hi');
    expect(cellValueAsPrimitive(42)).toBe(42);
    expect(cellValueAsPrimitive(true)).toBe(true);
    const d = new Date('2024-01-02T03:04:05Z');
    expect(cellValueAsPrimitive(d)).toBe(d);
  });

  it('rich-text → joined run text (string)', () => {
    const rt = makeRichText([{ text: 'Hello ' }, { text: 'world', font: { b: true } }]);
    expect(cellValueAsPrimitive({ kind: 'rich-text', runs: rt })).toBe('Hello world');
  });

  it('formula with cachedValue → cachedValue', () => {
    const f = { kind: 'formula' as const, t: 'normal' as const, formula: 'A1+1', cachedValue: 7 };
    expect(cellValueAsPrimitive(f)).toBe(7);
  });

  it('formula without cachedValue → null', () => {
    const f = { kind: 'formula' as const, t: 'normal' as const, formula: 'A1+1' };
    expect(cellValueAsPrimitive(f)).toBe(null);
  });

  it('error → error code (string)', () => {
    expect(cellValueAsPrimitive(makeErrorValue('#NAME?'))).toBe('#NAME?');
  });

  it('duration → ms (number)', () => {
    expect(cellValueAsPrimitive(makeDurationValue(1500))).toBe(1500);
  });
});
