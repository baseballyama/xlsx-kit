// Tests for borderToCss — Border value-object → CSS-property record.

import { describe, expect, it } from 'vitest';
import { borderToCss, makeBorder, makeSide } from '../../src/styles/borders';
import { makeColor } from '../../src/styles/colors';

describe('borderToCss', () => {
  it('returns {} for undefined / empty Border', () => {
    expect(borderToCss(undefined)).toEqual({});
    expect(borderToCss(makeBorder({}))).toEqual({});
  });

  it('emits all four edges for a uniform thin border with rgb colour', () => {
    const black = makeColor({ rgb: '000000' });
    const side = makeSide({ style: 'thin', color: black });
    const border = makeBorder({ top: side, right: side, bottom: side, left: side });
    expect(borderToCss(border)).toEqual({
      'border-top': '1px solid #000000',
      'border-right': '1px solid #000000',
      'border-bottom': '1px solid #000000',
      'border-left': '1px solid #000000',
    });
  });

  it('emits only the sides that are present and have a recognised style', () => {
    expect(
      borderToCss(
        makeBorder({
          top: makeSide({ style: 'medium', color: makeColor({ rgb: 'FF0000' }) }),
          // bottom present but no style → skipped
          bottom: makeSide({ color: makeColor({ rgb: '00FF00' }) }),
        }),
      ),
    ).toEqual({ 'border-top': '2px solid #FF0000' });
  });

  it('maps thick/double/dashed/dotted variants to the right CSS shorthands', () => {
    expect(
      borderToCss(
        makeBorder({
          top: makeSide({ style: 'thick' }),
          right: makeSide({ style: 'double' }),
          bottom: makeSide({ style: 'dashed' }),
          left: makeSide({ style: 'dotted' }),
        }),
      ),
    ).toEqual({
      'border-top': '3px solid currentColor',
      'border-right': '3px double currentColor',
      'border-bottom': '1px dashed currentColor',
      'border-left': '1px dotted currentColor',
    });
  });

  it('falls back to currentColor when side has no resolvable colour', () => {
    expect(borderToCss(makeBorder({ top: makeSide({ style: 'thin' }) }))).toEqual({
      'border-top': '1px solid currentColor',
    });
    // theme colour is unresolvable without a theme
    expect(
      borderToCss(makeBorder({ top: makeSide({ style: 'thin', color: makeColor({ theme: 1 }) }) })),
    ).toEqual({ 'border-top': '1px solid currentColor' });
  });

  it('skips diagonal / vertical / horizontal sides (no CSS equivalent)', () => {
    expect(
      borderToCss(
        makeBorder({
          diagonal: makeSide({ style: 'thin', color: makeColor({ rgb: '000000' }) }),
          vertical: makeSide({ style: 'thin' }),
          horizontal: makeSide({ style: 'thin' }),
        }),
      ),
    ).toEqual({});
  });
});
