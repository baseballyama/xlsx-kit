// Tests for fillToCss — Fill value-object → CSS-property record.

import { describe, expect, it } from 'vitest';
import { makeColor } from '../../src/styles/colors';
import {
  fillToCss,
  makeGradientFill,
  makeGradientStop,
  makePatternFill,
} from '../../src/styles/fills';

describe('fillToCss', () => {
  it('returns {} for undefined / "none" / no patternType', () => {
    expect(fillToCss(undefined)).toEqual({});
    expect(fillToCss(makePatternFill())).toEqual({});
    expect(fillToCss(makePatternFill({ patternType: 'none' }))).toEqual({});
  });

  it('renders solid pattern fill as background-color (alpha dropped)', () => {
    expect(
      fillToCss(makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: '00FF00' }) })),
    ).toEqual({ 'background-color': '#00FF00' });
  });

  it('skips solid fill with theme-only fgColor (cannot resolve without theme)', () => {
    expect(fillToCss(makePatternFill({ patternType: 'solid', fgColor: makeColor({ theme: 1 }) }))).toEqual({});
  });

  it('non-solid pattern collapses to bgColor (or fg fallback)', () => {
    expect(
      fillToCss(
        makePatternFill({
          patternType: 'darkGrid',
          bgColor: makeColor({ rgb: 'CCCCCC' }),
          fgColor: makeColor({ rgb: '000000' }),
        }),
      ),
    ).toEqual({ 'background-color': '#CCCCCC' });
    // missing bg → falls back to fg
    expect(
      fillToCss(
        makePatternFill({
          patternType: 'lightTrellis',
          fgColor: makeColor({ rgb: 'AAAAAA' }),
        }),
      ),
    ).toEqual({ 'background-color': '#AAAAAA' });
  });

  it('linear gradient emits linear-gradient(<deg>, stops…) with %-based positions', () => {
    const fill = makeGradientFill({
      type: 'linear',
      degree: 90,
      stops: [
        makeGradientStop(0, makeColor({ rgb: 'FF0000' })),
        makeGradientStop(1, makeColor({ rgb: '0000FF' })),
      ],
    });
    expect(fillToCss(fill)).toEqual({
      'background-image': 'linear-gradient(90deg, #FF0000 0.00%, #0000FF 100.00%)',
    });
  });

  it('path gradient emits radial-gradient(circle, stops…)', () => {
    const fill = makeGradientFill({
      type: 'path',
      stops: [
        makeGradientStop(0, makeColor({ rgb: 'FFFFFF' })),
        makeGradientStop(0.5, makeColor({ rgb: '888888' })),
        makeGradientStop(1, makeColor({ rgb: '000000' })),
      ],
    });
    expect(fillToCss(fill)).toEqual({
      'background-image': 'radial-gradient(circle, #FFFFFF 0.00%, #888888 50.00%, #000000 100.00%)',
    });
  });

  it('returns {} when no gradient stops resolve to hex (theme-only stops)', () => {
    const fill = makeGradientFill({
      type: 'linear',
      stops: [makeGradientStop(0, makeColor({ theme: 1 })), makeGradientStop(1, makeColor({ theme: 4 }))],
    });
    expect(fillToCss(fill)).toEqual({});
  });
});
