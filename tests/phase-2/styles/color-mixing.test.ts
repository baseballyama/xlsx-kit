// Tests for lighten / darken / mixColors.

import { describe, expect, it } from 'vitest';
import { darken, lighten, mixColors } from '../../../src/xlsx/styles/colors';

describe('lighten', () => {
  it('amount=0 returns the input unchanged', () => {
    expect(lighten('FF112233', 0)).toBe('FF112233');
  });

  it('amount=1 returns pure white (alpha preserved)', () => {
    expect(lighten('FF112233', 1)).toBe('FFFFFFFF');
    expect(lighten('80112233', 1)).toBe('80FFFFFF');
  });

  it('amount=0.5 mixes halfway', () => {
    // r=0x11=17 → 17 + (255-17)*0.5 = 136 = 0x88
    // g=0x22=34 → 34 + (255-34)*0.5 = 144.5 → round 145 = 0x91
    // b=0x33=51 → 51 + (255-51)*0.5 = 153 = 0x99
    expect(lighten('FF112233', 0.5)).toBe('FF889199');
  });

  it('clamps amount to [0, 1]', () => {
    expect(lighten('FF112233', -1)).toBe('FF112233');
    expect(lighten('FF112233', 2)).toBe('FFFFFFFF');
  });

  it('accepts 6-char input (alpha defaults to 00 per Excel convention)', () => {
    expect(lighten('112233', 0.5)).toBe('00889199');
  });
});

describe('darken', () => {
  it('amount=0 unchanged', () => {
    expect(darken('FFAABBCC', 0)).toBe('FFAABBCC');
  });

  it('amount=1 → black (alpha preserved)', () => {
    expect(darken('FFAABBCC', 1)).toBe('FF000000');
    expect(darken('80AABBCC', 1)).toBe('80000000');
  });

  it('amount=0.5 halves channel values', () => {
    // r=0xAA=170, g=0xBB=187, b=0xCC=204 → ×0.5 = 85, 93.5→94, 102
    expect(darken('FFAABBCC', 0.5)).toBe('FF555E66');
  });
});

describe('mixColors', () => {
  it('t=0 returns A; t=1 returns B', () => {
    expect(mixColors('FFFF0000', 'FF0000FF', 0)).toBe('FFFF0000');
    expect(mixColors('FFFF0000', 'FF0000FF', 1)).toBe('FF0000FF');
  });

  it('t=0.5 yields the channel-wise midpoint', () => {
    // (255+0)/2 = 128, 0, (0+255)/2 = 128 → "FF800080"
    expect(mixColors('FFFF0000', 'FF0000FF', 0.5)).toBe('FF800080');
  });

  it('alpha is interpolated too', () => {
    expect(mixColors('00000000', 'FF000000', 0.5)).toBe('80000000');
  });

  it('clamps t to [0, 1]', () => {
    expect(mixColors('FF000000', 'FFFFFFFF', -1)).toBe('FF000000');
    expect(mixColors('FF000000', 'FFFFFFFF', 2)).toBe('FFFFFFFF');
  });
});
