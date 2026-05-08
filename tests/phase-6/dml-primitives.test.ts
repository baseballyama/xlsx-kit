import { describe, expect, it } from 'vitest';
import {
  type ColorMod,
  type DmlColorWithMods,
  makeColor,
  makeSchemeColor,
  makeSrgbColor,
} from '../../src/xlsx/drawing/dml/colors';
import {
  parseDmlColor,
  parseFill,
  parseLine,
  parseShapeProperties,
  serializeDmlColor,
  serializeFill,
  serializeLine,
  serializeShapeProperties,
} from '../../src/xlsx/drawing/dml/dml-xml';
import {
  type Fill,
  makeGradientFill,
  makeNoFill,
  makePatternFill,
  makeSolidFill,
  PRESET_PATTERN_NAMES,
} from '../../src/xlsx/drawing/dml/fill';
import { type LineProperties, makeLine } from '../../src/xlsx/drawing/dml/line';
import { makeShapeProperties, type ShapeProperties } from '../../src/xlsx/drawing/dml/shape-properties';
import { parseXml } from '../../src/xml/parser';
import { findChild } from '../../src/xml/tree';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const NSDECL = `xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"`;

const wrap = (inner: string): string => `<c:spPr ${NSDECL}>${inner}</c:spPr>`;

const parseInWrap = (inner: string) => parseXml(wrap(inner));

const roundTripColor = (c: DmlColorWithMods): DmlColorWithMods => {
  const xml = `<c:wrap ${NSDECL}>${serializeDmlColor(c)}</c:wrap>`;
  const parsed = parseDmlColor(parseXml(xml));
  if (!parsed) throw new Error('color round-trip: parse failed');
  return parsed;
};

const roundTripFill = (f: Fill): Fill => {
  const xml = `<c:wrap ${NSDECL}>${serializeFill(f)}</c:wrap>`;
  const parsed = parseFill(parseXml(xml));
  if (!parsed) throw new Error('fill round-trip: parse failed');
  return parsed;
};

const roundTripLine = (ln: LineProperties): LineProperties => {
  const xml = `<c:wrap ${NSDECL}>${serializeLine(ln)}</c:wrap>`;
  const lnEl = findChild(parseXml(xml), `{${A_NS}}ln`);
  if (!lnEl) throw new Error('line round-trip: <a:ln> not found');
  return parseLine(lnEl);
};

const roundTripShapeProperties = (sp: ShapeProperties): ShapeProperties => {
  const xml = serializeShapeProperties(sp);
  const docXml = `<root ${NSDECL}>${xml}</root>`;
  const root = parseXml(docXml);
  const spEl = findChild(root, `{${C_NS}}spPr`);
  if (!spEl) throw new Error('shape props round-trip: <c:spPr> not found');
  return parseShapeProperties(spEl);
};

describe('DmlColor round-trip', () => {
  it('preserves srgbClr value', () => {
    expect(roundTripColor(makeColor(makeSrgbColor('FF8800')))).toEqual({
      base: { kind: 'srgb', value: 'FF8800' },
      mods: [],
    });
  });

  it('preserves schemeClr name', () => {
    expect(roundTripColor(makeColor(makeSchemeColor('accent3')))).toEqual({
      base: { kind: 'schemeClr', value: 'accent3' },
      mods: [],
    });
  });

  it('preserves modifier ordering — lumMod, lumOff, tint, alpha', () => {
    const mods: ColorMod[] = [
      { kind: 'lumMod', val: 75000 },
      { kind: 'lumOff', val: 25000 },
      { kind: 'tint', val: 50000 },
      { kind: 'alpha', val: 80000 },
    ];
    const back = roundTripColor(makeColor(makeSchemeColor('accent1'), mods));
    expect(back.mods).toEqual(mods);
  });

  it('preserves valueless mods (gray / inv)', () => {
    const back = roundTripColor(makeColor(makeSrgbColor('123456'), [{ kind: 'gray' }, { kind: 'inv' }]));
    expect(back.mods).toEqual([{ kind: 'gray' }, { kind: 'inv' }]);
  });

  it('preserves hslClr', () => {
    expect(roundTripColor({ base: { kind: 'hslClr', hue: 12000000, sat: 80000, lum: 40000 }, mods: [] })).toEqual({
      base: { kind: 'hslClr', hue: 12000000, sat: 80000, lum: 40000 },
      mods: [],
    });
  });

  it('preserves prstClr', () => {
    expect(roundTripColor({ base: { kind: 'prstClr', value: 'crimson' }, mods: [] })).toEqual({
      base: { kind: 'prstClr', value: 'crimson' },
      mods: [],
    });
  });
});

describe('Fill round-trip', () => {
  it('preserves noFill / grpFill sentinels', () => {
    expect(roundTripFill(makeNoFill())).toEqual({ kind: 'noFill' });
    expect(roundTripFill({ kind: 'grpFill' })).toEqual({ kind: 'grpFill' });
  });

  it('preserves solidFill with mods', () => {
    const fill = makeSolidFill(makeColor(makeSchemeColor('accent2'), [{ kind: 'lumMod', val: 60000 }]));
    expect(roundTripFill(fill)).toEqual(fill);
  });

  it('preserves gradFill with stops + linear direction', () => {
    const fill = makeGradientFill({
      flip: 'x',
      rotWithShape: true,
      stops: [
        { pos: 0, color: makeColor(makeSrgbColor('FF0000')) },
        { pos: 50000, color: makeColor(makeSrgbColor('00FF00')) },
        { pos: 100000, color: makeColor(makeSrgbColor('0000FF')) },
      ],
      lineDir: { kind: 'lin', ang: 5400000, scaled: true },
    });
    const back = roundTripFill(fill);
    if (back.kind !== 'gradFill') throw new Error('expected gradFill');
    expect(back.stops.length).toBe(3);
    expect(back.stops[1]?.pos).toBe(50000);
    expect(back.lineDir).toEqual({ kind: 'lin', ang: 5400000, scaled: true });
  });

  it('preserves gradFill with path direction', () => {
    const fill = makeGradientFill({
      stops: [{ pos: 0, color: makeColor(makeSrgbColor('000000')) }],
      lineDir: {
        kind: 'path',
        pathType: 'circle',
        tileRect: { l: 50000, t: 50000, r: 50000, b: 50000 },
      },
    });
    const back = roundTripFill(fill);
    if (back.kind !== 'gradFill') throw new Error('expected gradFill');
    expect(back.lineDir).toEqual({
      kind: 'path',
      pathType: 'circle',
      tileRect: { l: 50000, t: 50000, r: 50000, b: 50000 },
    });
  });

  it('preserves pattFill with foreground/background colors', () => {
    const fill = makePatternFill({
      preset: 'pct50',
      fgClr: makeColor(makeSrgbColor('AABBCC')),
      bgClr: makeColor(makeSrgbColor('FFFFFF')),
    });
    expect(roundTripFill(fill)).toEqual(fill);
  });

  it('preserves blipFill with embed rId + tile + srcRect', () => {
    const fill: Fill = {
      kind: 'blipFill',
      blip: { embedRId: 'rId7', cstate: 'screen' },
      tile: { tx: 0, ty: 0, sx: 100000, sy: 100000, flip: 'none', algn: 'tl' },
      srcRect: { l: 1000, t: 2000, r: 3000, b: 4000 },
      dpi: 96,
      rotWithShape: false,
    };
    const back = roundTripFill(fill);
    if (back.kind !== 'blipFill') throw new Error('expected blipFill');
    expect(back.blip.embedRId).toBe('rId7');
    expect(back.tile?.algn).toBe('tl');
    expect(back.srcRect).toEqual({ l: 1000, t: 2000, r: 3000, b: 4000 });
    expect(back.dpi).toBe(96);
    expect(back.rotWithShape).toBe(false);
  });

  it('lists ECMA-376 preset pattern names (54 entries, openpyxl-compatible)', () => {
    expect(PRESET_PATTERN_NAMES.length).toBe(54);
    expect(PRESET_PATTERN_NAMES).toContain('pct5');
    expect(PRESET_PATTERN_NAMES).toContain('zigZag');
    expect(new Set(PRESET_PATTERN_NAMES).size).toBe(54);
  });
});

describe('Line round-trip', () => {
  it('preserves width / cap / cmpd / algn + solidFill', () => {
    const ln = makeLine({
      w: 25400,
      cap: 'rnd',
      cmpd: 'thickThin',
      algn: 'ctr',
      fill: makeSolidFill(makeColor(makeSrgbColor('123456'))),
    });
    expect(roundTripLine(ln)).toEqual(ln);
  });

  it('preserves preset dash + miter join with limit', () => {
    const ln = makeLine({
      dash: { kind: 'preset', val: 'lgDashDot' },
      join: { kind: 'miter', lim: 200000 },
    });
    expect(roundTripLine(ln)).toEqual(ln);
  });

  it('preserves custom dash pattern (paired d/sp)', () => {
    const ln = makeLine({ dash: { kind: 'custDash', pattern: [400, 100, 200, 100] } });
    expect(roundTripLine(ln)).toEqual(ln);
  });

  it('preserves head/tail end markers', () => {
    const ln = makeLine({
      headEnd: { type: 'arrow', w: 'med', len: 'lg' },
      tailEnd: { type: 'triangle' },
    });
    expect(roundTripLine(ln)).toEqual(ln);
  });
});

describe('ShapeProperties round-trip', () => {
  it('preserves bwMode + xfrm + fill + ln', () => {
    const sp = makeShapeProperties({
      bwMode: 'auto',
      xfrm: {
        rot: 5400000,
        flipH: true,
        off: { x: 100, y: 200 },
        ext: { cx: 9144000, cy: 6858000 },
      },
      fill: makeSolidFill(makeColor(makeSchemeColor('accent4'), [{ kind: 'tint', val: 60000 }])),
      ln: makeLine({ w: 19050, cap: 'flat' }),
    });
    expect(roundTripShapeProperties(sp)).toEqual(sp);
  });

  it('emits bare wrapper when properties are absent', () => {
    expect(serializeShapeProperties({})).toBe('<c:spPr></c:spPr>');
    expect(roundTripShapeProperties({})).toEqual({});
  });

  it('uses caller-supplied wrapper tag', () => {
    expect(serializeShapeProperties({ fill: makeNoFill() }, 'a:spPr')).toBe('<a:spPr><a:noFill/></a:spPr>');
  });
});

describe('Fill in arbitrary parent works regardless of order', () => {
  it('finds the fill even when sibling elements precede it', () => {
    const inner = `<a:dummy/><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>`;
    const fill = parseFill(parseInWrap(inner));
    if (!fill || fill.kind !== 'solidFill') throw new Error('expected solidFill');
    expect(fill.color.base).toEqual({ kind: 'srgb', value: 'ABCDEF' });
  });
});
