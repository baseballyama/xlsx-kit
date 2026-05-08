// Tests for alignmentToCss — Alignment value-object → CSS-property record.

import { describe, expect, it } from 'vitest';
import { alignmentToCss, makeAlignment } from '../../../src/styles/alignment';

describe('alignmentToCss', () => {
  it('returns {} for undefined / empty Alignment', () => {
    expect(alignmentToCss(undefined)).toEqual({});
    expect(alignmentToCss(makeAlignment({}))).toEqual({});
  });

  it('maps horizontal values to text-align (general → unset)', () => {
    expect(alignmentToCss(makeAlignment({ horizontal: 'left' }))).toEqual({ 'text-align': 'left' });
    expect(alignmentToCss(makeAlignment({ horizontal: 'center' }))).toEqual({ 'text-align': 'center' });
    expect(alignmentToCss(makeAlignment({ horizontal: 'right' }))).toEqual({ 'text-align': 'right' });
    expect(alignmentToCss(makeAlignment({ horizontal: 'justify' }))).toEqual({ 'text-align': 'justify' });
    // 'general' has no CSS equivalent; caller decides per cell type.
    expect(alignmentToCss(makeAlignment({ horizontal: 'general' }))).toEqual({});
  });

  it('maps vertical values to vertical-align (table-cell semantics)', () => {
    expect(alignmentToCss(makeAlignment({ vertical: 'top' }))).toEqual({ 'vertical-align': 'top' });
    expect(alignmentToCss(makeAlignment({ vertical: 'center' }))).toEqual({ 'vertical-align': 'middle' });
    expect(alignmentToCss(makeAlignment({ vertical: 'bottom' }))).toEqual({ 'vertical-align': 'bottom' });
  });

  it('maps wrapText to white-space: pre-wrap', () => {
    expect(alignmentToCss(makeAlignment({ wrapText: true }))).toEqual({ 'white-space': 'pre-wrap' });
    expect(alignmentToCss(makeAlignment({ wrapText: false }))).toEqual({});
  });

  it('maps textRotation to a counter-rotation (Excel ccw → CSS cw); 0 omits, 255 = vertical-rl', () => {
    expect(alignmentToCss(makeAlignment({ textRotation: 90 }))).toEqual({
      transform: 'rotate(-90deg)',
      'transform-origin': 'center center',
    });
    expect(alignmentToCss(makeAlignment({ textRotation: 0 }))).toEqual({});
    expect(alignmentToCss(makeAlignment({ textRotation: 255 }))).toEqual({ 'writing-mode': 'vertical-rl' });
  });

  it('maps indent (>0) to padding-left in em; 0 is omitted', () => {
    expect(alignmentToCss(makeAlignment({ indent: 2 }))).toEqual({ 'padding-left': '2em' });
    expect(alignmentToCss(makeAlignment({ indent: 0 }))).toEqual({});
  });

  it('combines multiple properties into a single record', () => {
    expect(
      alignmentToCss(
        makeAlignment({
          horizontal: 'right',
          vertical: 'top',
          wrapText: true,
          indent: 1,
        }),
      ),
    ).toEqual({
      'text-align': 'right',
      'vertical-align': 'top',
      'white-space': 'pre-wrap',
      'padding-left': '1em',
    });
  });
});
