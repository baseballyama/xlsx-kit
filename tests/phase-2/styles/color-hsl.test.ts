// Tests for hexToHsl / hslToHex.

import { describe, expect, it } from 'vitest';
import { hexToHsl, hslToHex } from '../../../src/styles/colors';

describe('hexToHsl', () => {
  it('white → l=1, s=0', () => {
    const { h, s, l, a } = hexToHsl('FFFFFFFF');
    expect(l).toBeCloseTo(1, 6);
    expect(s).toBeCloseTo(0, 6);
    expect(a).toBe(255);
    expect(h).toBe(0);
  });

  it('black → l=0, s=0', () => {
    const { l, s } = hexToHsl('FF000000');
    expect(l).toBeCloseTo(0, 6);
    expect(s).toBeCloseTo(0, 6);
  });

  it('pure red → h=0, s=1, l=0.5', () => {
    const { h, s, l } = hexToHsl('FFFF0000');
    expect(h).toBeCloseTo(0, 4);
    expect(s).toBeCloseTo(1, 6);
    expect(l).toBeCloseTo(0.5, 6);
  });

  it('pure green → h=120', () => {
    expect(hexToHsl('FF00FF00').h).toBeCloseTo(120, 4);
  });

  it('pure blue → h=240', () => {
    expect(hexToHsl('FF0000FF').h).toBeCloseTo(240, 4);
  });

  it('alpha is preserved as a byte', () => {
    expect(hexToHsl('80FF0000').a).toBe(0x80);
  });
});

describe('hslToHex', () => {
  it('round-trips pure red', () => {
    expect(hslToHex(0, 1, 0.5)).toBe('FFFF0000');
  });

  it('round-trips pure green', () => {
    expect(hslToHex(120, 1, 0.5)).toBe('FF00FF00');
  });

  it('round-trips pure blue', () => {
    expect(hslToHex(240, 1, 0.5)).toBe('FF0000FF');
  });

  it('white = l=1', () => {
    expect(hslToHex(0, 0, 1)).toBe('FFFFFFFF');
  });

  it('h wraps mod-360', () => {
    expect(hslToHex(360, 1, 0.5)).toBe(hslToHex(0, 1, 0.5));
    expect(hslToHex(-120, 1, 0.5)).toBe(hslToHex(240, 1, 0.5));
  });

  it('alpha argument flows to the high byte', () => {
    expect(hslToHex(0, 1, 0.5, 0)).toBe('00FF0000');
    expect(hslToHex(0, 1, 0.5, 0x80)).toBe('80FF0000');
  });
});

describe('hex ↔ HSL round-trip', () => {
  it('arbitrary hex round-trips within rounding tolerance', () => {
    const original = 'FF305496';
    const { h, s, l, a } = hexToHsl(original);
    const back = hslToHex(h, s, l, a);
    expect(back).toBe(original);
  });
});
