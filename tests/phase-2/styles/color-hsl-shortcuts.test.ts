// Tests for rotateHue / adjustSaturation / adjustLightness.

import { describe, expect, it } from 'vitest';
import {
  adjustLightness,
  adjustSaturation,
  hexToHsl,
  rotateHue,
} from '../../../src/styles/colors';

describe('rotateHue', () => {
  it('rotating red by 120° → green', () => {
    expect(rotateHue('FFFF0000', 120)).toBe('FF00FF00');
  });

  it('rotating red by -120° (or +240°) → blue', () => {
    expect(rotateHue('FFFF0000', -120)).toBe('FF0000FF');
  });

  it('360° rotation → identity', () => {
    expect(rotateHue('FF305496', 360)).toBe('FF305496');
  });

  it('preserves alpha', () => {
    expect(rotateHue('80FF0000', 120)).toBe('8000FF00');
  });
});

describe('adjustSaturation', () => {
  it('+1 from a fully saturated color clamps to 1', () => {
    const before = hexToHsl('FFFF0000');
    const after = hexToHsl(adjustSaturation('FFFF0000', 1));
    expect(after.s).toBeCloseTo(1);
    expect(after.h).toBeCloseTo(before.h);
  });

  it('-1 fully desaturates', () => {
    const after = hexToHsl(adjustSaturation('FFFF0000', -1));
    expect(after.s).toBeCloseTo(0);
  });

  it('preserves alpha', () => {
    expect(adjustSaturation('80FF0000', -1)).toMatch(/^80/);
  });
});

describe('adjustLightness', () => {
  it('+0.5 from black → mid-gray', () => {
    const after = hexToHsl(adjustLightness('FF000000', 0.5));
    expect(after.l).toBeCloseTo(0.5);
    expect(after.s).toBeCloseTo(0);
  });

  it('-1 from any color → black', () => {
    expect(adjustLightness('FFFF0000', -1)).toBe('FF000000');
  });

  it('+1 from any color → white', () => {
    expect(adjustLightness('FFFF0000', 1)).toBe('FFFFFFFF');
  });

  it('preserves alpha', () => {
    expect(adjustLightness('80112233', 0.2)).toMatch(/^80/);
  });
});
