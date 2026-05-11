// Tests for colorToHex — Color value-object → ARGB hex readback.

import { describe, expect, it } from 'vitest';
import { colorToHex, makeColor } from '../../src/styles/colors';

describe('colorToHex', () => {
  it('returns the normalised ARGB hex when rgb is set', () => {
    expect(colorToHex(makeColor({ rgb: 'FF0000' }))).toBe('00FF0000');
    expect(colorToHex(makeColor({ rgb: 'FFAABBCC' }))).toBe('FFAABBCC');
  });

  it('resolves indexed colours via the legacy palette', () => {
    // Index 2 = '00FF0000' (pure red) per COLOR_INDEX.
    expect(colorToHex(makeColor({ indexed: 2 }))).toBe('00FF0000');
  });

  it('returns undefined for theme-only colours (cannot resolve without theme)', () => {
    expect(colorToHex(makeColor({ theme: 1 }))).toBeUndefined();
  });

  it('returns undefined for auto colours', () => {
    expect(colorToHex(makeColor({ auto: true }))).toBeUndefined();
  });

  it('returns undefined for empty Color and undefined input', () => {
    expect(colorToHex(makeColor({}))).toBeUndefined();
    expect(colorToHex(undefined)).toBeUndefined();
  });

  it('prefers rgb over indexed when both are set', () => {
    // makeColor accepts both for forward compat; rgb wins on readback.
    expect(colorToHex(makeColor({ rgb: '0000FF', indexed: 2 }))).toBe('000000FF');
  });
});
