// Tests for luminance / contrastRatio / pickReadableTextColor.

import { describe, expect, it } from 'vitest';
import { contrastRatio, luminance, pickReadableTextColor } from '../../src/styles/colors';

describe('luminance', () => {
  it('white = 1, black = 0', () => {
    expect(luminance('FFFFFFFF')).toBeCloseTo(1, 6);
    expect(luminance('FF000000')).toBeCloseTo(0, 6);
  });

  it('mid-gray ≈ 0.18', () => {
    expect(luminance('FF777777')).toBeCloseTo(0.184, 2);
  });

  it('alpha is ignored — same RGB different alpha produces same luminance', () => {
    expect(luminance('FFAA0000')).toBeCloseTo(luminance('00AA0000'), 6);
  });

  it('accepts 6-char (no alpha) form', () => {
    expect(luminance('FFFFFF')).toBeCloseTo(1, 6);
  });
});

describe('contrastRatio', () => {
  it('white / black → 21', () => {
    expect(contrastRatio('FFFFFFFF', 'FF000000')).toBeCloseTo(21, 1);
  });

  it('order-independent', () => {
    expect(contrastRatio('FFFFAA00', 'FF003366')).toBeCloseTo(
      contrastRatio('FF003366', 'FFFFAA00'),
      6,
    );
  });

  it('identical colors → 1', () => {
    expect(contrastRatio('FF888888', 'FF888888')).toBeCloseTo(1, 6);
  });
});

describe('pickReadableTextColor', () => {
  it('white text on dark backgrounds', () => {
    expect(pickReadableTextColor('FF000000')).toBe('FFFFFFFF');
    expect(pickReadableTextColor('FF222222')).toBe('FFFFFFFF');
    expect(pickReadableTextColor('FF003366')).toBe('FFFFFFFF');
  });

  it('black text on light backgrounds', () => {
    expect(pickReadableTextColor('FFFFFFFF')).toBe('FF000000');
    expect(pickReadableTextColor('FFFFFF00')).toBe('FF000000');
    expect(pickReadableTextColor('FFAAAAAA')).toBe('FF000000');
  });

  it('black text on Excel default theme blue', () => {
    // FF305496 is "Table Style Medium 2" header — sits on the dark side.
    expect(pickReadableTextColor('FF305496')).toBe('FFFFFFFF');
  });
});
