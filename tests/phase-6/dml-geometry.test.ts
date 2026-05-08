import { describe, expect, it } from 'vitest';
import {
  parseGeometry,
  parseShapeProperties,
  serializeGeometry,
  serializeShapeProperties,
} from '../../src/xlsx/drawing/dml/dml-xml';
import {
  type CustomGeometry,
  type Geometry,
  isPresetShapeName,
  makeCustomGeometry,
  makePresetGeometry,
  PRESET_SHAPE_NAMES,
} from '../../src/xlsx/drawing/dml/geometry';
import { makeShapeProperties, type ShapeProperties } from '../../src/xlsx/drawing/dml/shape-properties';
import { parseXml } from '../../src/xml/parser';
import { findChild } from '../../src/xml/tree';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NSDECL = `xmlns:c="${C_NS}" xmlns:a="${A_NS}"`;

const roundTripGeometry = (g: Geometry): Geometry => {
  const xml = `<c:wrap ${NSDECL}>${serializeGeometry(g)}</c:wrap>`;
  const parsed = parseGeometry(parseXml(xml));
  if (!parsed) throw new Error('geometry round-trip: parse failed');
  return parsed;
};

const roundTripShapeProperties = (sp: ShapeProperties): ShapeProperties => {
  const xml = `<root ${NSDECL}>${serializeShapeProperties(sp)}</root>`;
  const root = parseXml(xml);
  const spEl = findChild(root, `{${C_NS}}spPr`);
  if (!spEl) throw new Error('spPr round-trip: <c:spPr> missing');
  return parseShapeProperties(spEl);
};

describe('Preset shape catalogue', () => {
  it('lists ECMA-376 preset shapes (187 entries, deduped, openpyxl-compatible)', () => {
    expect(PRESET_SHAPE_NAMES.length).toBe(187);
    expect(new Set(PRESET_SHAPE_NAMES).size).toBe(187);
  });

  it('isPresetShapeName accepts canonical names + rejects unknowns', () => {
    expect(isPresetShapeName('rect')).toBe(true);
    expect(isPresetShapeName('roundRect')).toBe(true);
    expect(isPresetShapeName('flowChartProcess')).toBe(true);
    expect(isPresetShapeName('mathPlus')).toBe(true);
    expect(isPresetShapeName('actionButtonReturn')).toBe(true);
    expect(isPresetShapeName('chartPlus')).toBe(true);
    expect(isPresetShapeName('not-a-shape')).toBe(false);
  });
});

describe('PresetGeometry round-trip', () => {
  it('preserves prst with no avLst', () => {
    expect(roundTripGeometry(makePresetGeometry('rect'))).toEqual({ kind: 'preset', prst: 'rect' });
  });

  it('preserves prst with avLst guides', () => {
    const g = makePresetGeometry('roundRect', [{ name: 'adj', fmla: 'val 16667' }]);
    expect(roundTripGeometry(g)).toEqual(g);
  });

  it('round-trips every preset shape name through serializer + parser', () => {
    for (const name of PRESET_SHAPE_NAMES) {
      const back = roundTripGeometry(makePresetGeometry(name));
      if (back.kind !== 'preset') throw new Error('expected preset kind');
      expect(back.prst).toBe(name);
    }
  });
});

describe('CustomGeometry round-trip', () => {
  it('preserves a single triangle path (moveTo + 2× lnTo + close)', () => {
    const g: CustomGeometry = makeCustomGeometry({
      pathLst: [
        {
          w: 100,
          h: 100,
          fill: 'norm',
          stroke: true,
          extrusionOk: false,
          commands: [
            { kind: 'moveTo', pt: { x: 0, y: 100 } },
            { kind: 'lnTo', pt: { x: 50, y: 0 } },
            { kind: 'lnTo', pt: { x: 100, y: 100 } },
            { kind: 'close' },
          ],
        },
      ],
    });
    expect(roundTripGeometry(g)).toEqual(g);
  });

  it('preserves arcTo + cubicBezTo + quadBezTo commands', () => {
    const g: CustomGeometry = makeCustomGeometry({
      pathLst: [
        {
          commands: [
            { kind: 'moveTo', pt: { x: 0, y: 0 } },
            { kind: 'arcTo', wR: '50', hR: '50', stAng: '0', swAng: '5400000' },
            {
              kind: 'cubicBezTo',
              pts: [
                { x: 100, y: 0 },
                { x: 200, y: 50 },
                { x: 300, y: 100 },
              ],
            },
            {
              kind: 'quadBezTo',
              pts: [
                { x: 350, y: 200 },
                { x: 400, y: 300 },
              ],
            },
            { kind: 'close' },
          ],
        },
      ],
    });
    expect(roundTripGeometry(g)).toEqual(g);
  });

  it('preserves avLst / gdLst / ahLst / cxnLst / rect', () => {
    const g: CustomGeometry = makeCustomGeometry({
      avLst: [{ name: 'adj1', fmla: 'val 25000' }],
      gdLst: [
        { name: 'w2', fmla: '*/ w 1 2' },
        { name: 'h2', fmla: '*/ h 1 2' },
      ],
      ahLst: [
        {
          kind: 'xy',
          pos: { x: 'adj1', y: 'h2' },
          gdRefX: 'adj1',
          minX: '0',
          maxX: '50000',
        },
        {
          kind: 'polar',
          pos: { x: 'w2', y: 'h2' },
          gdRefAng: 'angAdj',
          minAng: '0',
          maxAng: '21600000',
        },
      ],
      cxnLst: [
        { ang: '0', pos: { x: '0', y: 'h2' } },
        { ang: '5400000', pos: { x: 'w2', y: '0' } },
      ],
      rect: { l: '0', t: '0', r: 'w2', b: 'h2' },
      pathLst: [
        {
          commands: [{ kind: 'moveTo', pt: { x: 0, y: 0 } }, { kind: 'close' }],
        },
      ],
    });
    const back = roundTripGeometry(g);
    if (back.kind !== 'custom') throw new Error('expected custom kind');
    expect(back.avLst).toEqual(g.avLst);
    expect(back.gdLst).toEqual(g.gdLst);
    expect(back.ahLst).toEqual(g.ahLst);
    expect(back.cxnLst).toEqual(g.cxnLst);
    expect(back.rect).toEqual(g.rect);
  });

  it('preserves multiple path entries', () => {
    const g: CustomGeometry = makeCustomGeometry({
      pathLst: [
        {
          commands: [{ kind: 'moveTo', pt: { x: 0, y: 0 } }, { kind: 'lnTo', pt: { x: 10, y: 10 } }, { kind: 'close' }],
        },
        {
          commands: [
            { kind: 'moveTo', pt: { x: 20, y: 20 } },
            { kind: 'lnTo', pt: { x: 30, y: 30 } },
            { kind: 'close' },
          ],
        },
      ],
    });
    const back = roundTripGeometry(g);
    if (back.kind !== 'custom') throw new Error('expected custom kind');
    expect(back.pathLst.length).toBe(2);
    expect(back.pathLst[1]?.commands[1]).toEqual({ kind: 'lnTo', pt: { x: 30, y: 30 } });
  });
});

describe('ShapeProperties carries geometry', () => {
  it('emits prstGeom inside spPr and round-trips it', () => {
    const sp = makeShapeProperties({ geometry: makePresetGeometry('ellipse') });
    expect(roundTripShapeProperties(sp)).toEqual(sp);
  });

  it('places geometry between xfrm and fill/ln (ECMA-376 element order)', () => {
    const sp = makeShapeProperties({
      xfrm: { off: { x: 100, y: 200 }, ext: { cx: 1000, cy: 2000 } },
      geometry: makePresetGeometry('rect'),
    });
    const xml = serializeShapeProperties(sp);
    const xfrmIdx = xml.indexOf('<a:xfrm');
    const geomIdx = xml.indexOf('<a:prstGeom');
    expect(xfrmIdx).toBeGreaterThan(-1);
    expect(geomIdx).toBeGreaterThan(xfrmIdx);
  });
});
