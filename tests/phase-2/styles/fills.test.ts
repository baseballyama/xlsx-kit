import { describe, expect, it } from 'vitest';
import { makeColor } from '../../../src/styles/colors';
import {
  DEFAULT_EMPTY_FILL,
  DEFAULT_GRAY_FILL,
  type Fill,
  makeFill,
  makeGradientFill,
  makeGradientStop,
  makePatternFill,
  PATTERN_TYPES,
} from '../../../src/styles/fills';
import { fillFromTree, fillToTree } from '../../../src/styles/fills.schema';
import { OpenXmlSchemaError } from '../../../src/utils/exceptions';
import { parseXml } from '../../../src/xml/parser';
import { serializeXml } from '../../../src/xml/serializer';

describe('PatternFill', () => {
  it('makePatternFill freezes the result', () => {
    const f = makePatternFill({ patternType: 'solid' });
    expect(Object.isFrozen(f)).toBe(true);
    expect(f.kind).toBe('pattern');
  });

  it('omits unset fields', () => {
    const f = makePatternFill();
    expect(Object.keys(f).sort()).toEqual(['kind']);
  });

  it('accepts every legal patternType', () => {
    for (const pt of PATTERN_TYPES) {
      expect(makePatternFill({ patternType: pt }).patternType).toBe(pt);
    }
  });

  it('rejects an unknown patternType', () => {
    // biome-ignore lint/suspicious/noExplicitAny: deliberately bad input
    expect(() => makePatternFill({ patternType: 'starfield' as any })).toThrowError(OpenXmlSchemaError);
  });

  it('freezes nested fgColor / bgColor on construction', () => {
    const f = makePatternFill({ patternType: 'solid', fgColor: { rgb: 'FF0000' }, bgColor: { rgb: '00FF00' } });
    expect(f.fgColor && Object.isFrozen(f.fgColor)).toBe(true);
    expect(f.bgColor && Object.isFrozen(f.bgColor)).toBe(true);
  });
});

describe('GradientFill', () => {
  it('defaults type to "linear" and stops to []', () => {
    const f = makeGradientFill();
    expect(f.kind).toBe('gradient');
    expect(f.type).toBe('linear');
    expect(f.stops).toEqual([]);
  });

  it('rejects an out-of-range stop position', () => {
    expect(() => makeGradientStop(-0.1, makeColor({ rgb: 'FF0000' }))).toThrowError(OpenXmlSchemaError);
    expect(() => makeGradientStop(1.5, makeColor({ rgb: 'FF0000' }))).toThrowError(OpenXmlSchemaError);
  });

  it('rejects an unknown gradient type', () => {
    // biome-ignore lint/suspicious/noExplicitAny: deliberately bad input
    expect(() => makeGradientFill({ type: 'radial' as any })).toThrowError(OpenXmlSchemaError);
  });

  it('freezes the stops array and each stop', () => {
    const f = makeGradientFill({
      stops: [makeGradientStop(0, makeColor({ rgb: 'FF0000' })), makeGradientStop(1, makeColor({ rgb: '0000FF' }))],
    });
    expect(Object.isFrozen(f.stops)).toBe(true);
    expect(f.stops[0] && Object.isFrozen(f.stops[0])).toBe(true);
  });
});

describe('makeFill / DEFAULT_*_FILL', () => {
  it('dispatches on kind', () => {
    const p = makeFill({ kind: 'pattern', patternType: 'solid' });
    expect(p.kind).toBe('pattern');
    const g = makeFill({ kind: 'gradient', type: 'linear' });
    expect(g.kind).toBe('gradient');
  });

  it('defaults to PatternFill when kind is omitted', () => {
    const f = makeFill({});
    expect(f.kind).toBe('pattern');
  });

  it('DEFAULT_EMPTY_FILL is the empty PatternFill', () => {
    expect(DEFAULT_EMPTY_FILL).toEqual({ kind: 'pattern' });
    expect(Object.isFrozen(DEFAULT_EMPTY_FILL)).toBe(true);
  });

  it('DEFAULT_GRAY_FILL is gray125', () => {
    expect(DEFAULT_GRAY_FILL).toEqual({ kind: 'pattern', patternType: 'gray125' });
  });
});

describe('Fill XML round-trip via fillToTree / fillFromTree', () => {
  const cases: ReadonlyArray<readonly [string, Fill]> = [
    ['empty PatternFill', makePatternFill()],
    ['solid PatternFill (no colours)', makePatternFill({ patternType: 'solid' })],
    [
      'solid PatternFill with fg+bg',
      makePatternFill({
        patternType: 'solid',
        fgColor: makeColor({ rgb: 'FFFF0000' }),
        bgColor: makeColor({ rgb: 'FF00FF00' }),
      }),
    ],
    [
      'lightGrid PatternFill with fg only',
      makePatternFill({ patternType: 'lightGrid', fgColor: makeColor({ theme: 1 }) }),
    ],
    ['linear gradient (default)', makeGradientFill()],
    [
      'linear gradient with two stops',
      makeGradientFill({
        type: 'linear',
        degree: 90,
        stops: [makeGradientStop(0, makeColor({ rgb: 'FF0000' })), makeGradientStop(1, makeColor({ rgb: '0000FF' }))],
      }),
    ],
    [
      'path gradient with insets',
      makeGradientFill({
        type: 'path',
        left: 0.1,
        right: 0.1,
        top: 0.2,
        bottom: 0.2,
        stops: [makeGradientStop(0, makeColor({ rgb: 'FFFFFF' })), makeGradientStop(1, makeColor({ rgb: '000000' }))],
      }),
    ],
  ];

  it.each(cases)('round-trips %s', (_label, fill) => {
    const xml = serializeXml(fillToTree(fill));
    const back = fillFromTree(parseXml(xml));
    expect(back).toEqual(fill);
  });

  it('emits the <fill> wrapper element', () => {
    const xml = new TextDecoder().decode(serializeXml(fillToTree(makePatternFill({ patternType: 'solid' }))));
    // <fill> is the root here, so it carries the xmlns declaration.
    expect(xml).toMatch(/<fill xmlns="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main">/);
    expect(xml).toContain('<patternFill patternType="solid"/>');
    expect(xml).toContain('</fill>');
  });

  it('rejects malformed wrapper input', () => {
    const noFill = parseXml(
      '<?xml version="1.0"?><other xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
    );
    expect(() => fillFromTree(noFill)).toThrowError(OpenXmlSchemaError);
    const emptyFill = parseXml(
      '<?xml version="1.0"?><fill xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
    );
    expect(() => fillFromTree(emptyFill)).toThrowError(OpenXmlSchemaError);
  });
});
