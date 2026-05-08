// Tests for fontToCss — Font value-object → CSS-property record.

import { describe, expect, it } from 'vitest';
import { makeColor } from '../../../src/xlsx/styles/colors';
import { fontToCss, makeFont } from '../../../src/xlsx/styles/fonts';

describe('fontToCss', () => {
  it('returns an empty record for empty / undefined Font (caller decides default)', () => {
    expect(fontToCss(makeFont({}))).toEqual({});
    expect(fontToCss(undefined)).toEqual({});
  });

  it('emits font-family + font-size in pt', () => {
    expect(fontToCss(makeFont({ name: 'Calibri', size: 11 }))).toEqual({
      'font-family': "'Calibri'",
      'font-size': '11pt',
    });
  });

  it('maps bold/italic/strike to weight/style/decoration', () => {
    expect(fontToCss(makeFont({ bold: true }))).toEqual({ 'font-weight': 'bold' });
    expect(fontToCss(makeFont({ italic: true }))).toEqual({ 'font-style': 'italic' });
    expect(fontToCss(makeFont({ strike: true }))).toEqual({ 'text-decoration': 'line-through' });
  });

  it('maps underline (any style) to text-decoration: underline; strike+underline combine', () => {
    expect(fontToCss(makeFont({ underline: 'single' }))).toEqual({ 'text-decoration': 'underline' });
    expect(fontToCss(makeFont({ underline: 'double', strike: true }))).toEqual({
      'text-decoration': 'underline line-through',
    });
  });

  it('renders rgb color as #RRGGBB (alpha dropped); skips theme-only colours', () => {
    expect(fontToCss(makeFont({ color: makeColor({ rgb: 'FF0000' }) }))).toEqual({ color: '#FF0000' });
    // theme/auto colours can't be resolved without a theme — caller cascades.
    expect(fontToCss(makeFont({ color: makeColor({ theme: 1 }) }))).toEqual({});
  });

  it('handles superscript/subscript: vertical-align + 0.83em fallback when size unset', () => {
    expect(fontToCss(makeFont({ vertAlign: 'superscript' }))).toEqual({
      'vertical-align': 'superscript',
      'font-size': '0.83em',
    });
    // explicit size wins
    expect(fontToCss(makeFont({ vertAlign: 'subscript', size: 9 }))).toEqual({
      'vertical-align': 'subscript',
      'font-size': '9pt',
    });
  });

  it('escapes single quotes in font-family names', () => {
    expect(fontToCss(makeFont({ name: "Comic Sans 'Pro'" }))['font-family']).toBe("'Comic Sans \\'Pro\\''");
  });
});
