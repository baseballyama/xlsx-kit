import { describe, expect, it } from 'vitest';
import { makeColor, makeSchemeColor, makeSrgbColor } from '../../src/xlsx/drawing/dml/colors';
import {
  parseEffects,
  parseShapeProperties,
  serializeEffects,
  serializeShapeProperties,
} from '../../src/xlsx/drawing/dml/dml-xml';
import {
  type Effect,
  type EffectContainer,
  type EffectsRef,
  makeEffectContainer,
  makeEffectList,
  PRESET_SHADOW_NAMES,
} from '../../src/xlsx/drawing/dml/effect';
import { makeNoFill, makeSolidFill } from '../../src/xlsx/drawing/dml/fill';
import { makeShapeProperties, type ShapeProperties } from '../../src/xlsx/drawing/dml/shape-properties';
import { parseXml } from '../../src/xml/parser';
import { findChild } from '../../src/xml/tree';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NSDECL = `xmlns:c="${C_NS}" xmlns:a="${A_NS}"`;

const roundTripEffects = (e: EffectsRef): EffectsRef => {
  const xml = `<c:wrap ${NSDECL}>${serializeEffects(e)}</c:wrap>`;
  const parsed = parseEffects(parseXml(xml));
  if (!parsed) throw new Error('effects round-trip: parse failed');
  return parsed;
};

const roundTripShapeProperties = (sp: ShapeProperties): ShapeProperties => {
  const xml = `<root ${NSDECL}>${serializeShapeProperties(sp)}</root>`;
  const root = parseXml(xml);
  const spEl = findChild(root, `{${C_NS}}spPr`);
  if (!spEl) throw new Error('spPr round-trip: <c:spPr> missing');
  return parseShapeProperties(spEl);
};

describe('Preset shadow catalogue', () => {
  it('lists 20 ECMA-376 preset shadows', () => {
    expect(PRESET_SHADOW_NAMES.length).toBe(20);
    expect(PRESET_SHADOW_NAMES[0]).toBe('shdw1');
    expect(PRESET_SHADOW_NAMES[19]).toBe('shdw20');
  });
});

describe('EffectList round-trip', () => {
  it('preserves blur (rad + grow)', () => {
    const e: EffectsRef = { kind: 'lst', list: makeEffectList([{ kind: 'blur', rad: 12700, grow: true }]) };
    expect(roundTripEffects(e)).toEqual(e);
  });

  it('preserves glow with color', () => {
    const e: EffectsRef = {
      kind: 'lst',
      list: makeEffectList([
        { kind: 'glow', rad: 38100, color: makeColor(makeSrgbColor('FF0000'), [{ kind: 'alpha', val: 60000 }]) },
      ]),
    };
    expect(roundTripEffects(e)).toEqual(e);
  });

  it('preserves outerShdw with full attribute set', () => {
    const e: EffectsRef = {
      kind: 'lst',
      list: makeEffectList([
        {
          kind: 'outerShdw',
          blurRad: 50800,
          dist: 38100,
          dir: 2700000,
          sx: 100000,
          sy: 100000,
          kx: 0,
          ky: 0,
          algn: 'tl',
          rotWithShape: false,
          color: makeColor(makeSrgbColor('000000'), [{ kind: 'alpha', val: 50000 }]),
        },
      ]),
    };
    expect(roundTripEffects(e)).toEqual(e);
  });

  it('preserves innerShdw + softEdge composed', () => {
    const e: EffectsRef = {
      kind: 'lst',
      list: makeEffectList([
        { kind: 'innerShdw', blurRad: 25400, dist: 12700, dir: 5400000, color: makeColor(makeSrgbColor('333333')) },
        { kind: 'softEdge', rad: 25400 },
      ]),
    };
    const back = roundTripEffects(e);
    if (back.kind !== 'lst') throw new Error('expected lst');
    expect(back.list.list.length).toBe(2);
    expect(back.list.list[0]).toEqual(e.list.list[0]);
    expect(back.list.list[1]).toEqual({ kind: 'softEdge', rad: 25400 });
  });

  it('preserves prstShdw with valid preset name', () => {
    const e: EffectsRef = {
      kind: 'lst',
      list: makeEffectList([
        { kind: 'prstShdw', prst: 'shdw17', dist: 38100, dir: 2700000, color: makeColor(makeSchemeColor('accent1')) },
      ]),
    };
    expect(roundTripEffects(e)).toEqual(e);
  });

  it('preserves reflection with full attribute set', () => {
    const e: EffectsRef = {
      kind: 'lst',
      list: makeEffectList([
        {
          kind: 'reflection',
          blurRad: 6350,
          stA: 50000,
          stPos: 0,
          endA: 300,
          endPos: 50000,
          dist: 5000,
          dir: 5400000,
          fadeDir: 5400000,
          sx: 100000,
          sy: -100000,
          kx: 0,
          ky: 0,
          algn: 'bl',
          rotWithShape: false,
        },
      ]),
    };
    expect(roundTripEffects(e)).toEqual(e);
  });

  it('preserves fillOverlay with nested fill', () => {
    const e: EffectsRef = {
      kind: 'lst',
      list: makeEffectList([
        { kind: 'fillOverlay', blend: 'mult', fill: makeSolidFill(makeColor(makeSrgbColor('AABBCC'))) },
      ]),
    };
    const back = roundTripEffects(e);
    if (back.kind !== 'lst') throw new Error('expected lst');
    const eff = back.list.list[0];
    if (!eff || eff.kind !== 'fillOverlay') throw new Error('expected fillOverlay');
    expect(eff.blend).toBe('mult');
    expect(eff.fill).toEqual(makeSolidFill(makeColor(makeSrgbColor('AABBCC'))));
  });

  it('preserves order of multi-leaf list', () => {
    const leaves: Effect[] = [
      { kind: 'blur', rad: 1000 },
      { kind: 'glow', rad: 2000, color: makeColor(makeSrgbColor('123456')) },
      { kind: 'softEdge', rad: 3000 },
    ];
    const back = roundTripEffects({ kind: 'lst', list: makeEffectList(leaves) });
    if (back.kind !== 'lst') throw new Error('expected lst');
    expect(back.list.list.map((l) => l.kind)).toEqual(['blur', 'glow', 'softEdge']);
  });
});

describe('EffectDag round-trip', () => {
  it('preserves a single cont with a blur leaf', () => {
    const e: EffectsRef = {
      kind: 'dag',
      children: [makeEffectContainer('sib', [{ kind: 'blur', rad: 12700 }])],
    };
    expect(roundTripEffects(e)).toEqual(e);
  });

  it('preserves nested cont (tree → sib → leaf)', () => {
    const e: EffectsRef = {
      kind: 'dag',
      children: [
        makeEffectContainer(
          'tree',
          [
            makeEffectContainer('sib', [
              { kind: 'glow', rad: 5000, color: makeColor(makeSrgbColor('FFFFFF')) },
              { kind: 'softEdge', rad: 1000 },
            ]),
          ],
          'outer',
        ),
      ],
    };
    const back = roundTripEffects(e);
    if (back.kind !== 'dag') throw new Error('expected dag');
    expect(back.children.length).toBe(1);
    const top = back.children[0] as EffectContainer;
    expect(top.type).toBe('tree');
    expect(top.name).toBe('outer');
    const inner = top.children[0] as EffectContainer;
    expect(inner.type).toBe('sib');
    expect(inner.children.length).toBe(2);
  });

  it('preserves leaf siblings at the dag root', () => {
    const e: EffectsRef = {
      kind: 'dag',
      children: [{ kind: 'blur', rad: 100 }, makeEffectContainer('sib', [{ kind: 'softEdge', rad: 200 }])],
    };
    const back = roundTripEffects(e);
    if (back.kind !== 'dag') throw new Error('expected dag');
    expect(back.children.length).toBe(2);
  });
});

describe('ShapeProperties carries effects', () => {
  it('emits effectLst inside spPr in ECMA order (after fill/ln)', () => {
    const sp = makeShapeProperties({
      fill: makeNoFill(),
      effects: { kind: 'lst', list: makeEffectList([{ kind: 'blur', rad: 1000 }]) },
    });
    const xml = serializeShapeProperties(sp);
    const fillIdx = xml.indexOf('<a:noFill');
    const effectIdx = xml.indexOf('<a:effectLst');
    expect(effectIdx).toBeGreaterThan(fillIdx);
    expect(roundTripShapeProperties(sp)).toEqual(sp);
  });
});
