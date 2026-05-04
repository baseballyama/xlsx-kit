import { describe, expect, it } from 'vitest';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart';
import {
  parseUserShapesXml,
  serializeUserShapes,
  userShapesToBytes,
} from '../../src/chart/user-shapes-xml';
import {
  type ChartDrawing,
  makeAbsSizeAnchor,
  makeChartDrawing,
  makeChartShape,
  makeRelSizeAnchor,
} from '../../src/chart/user-shapes';
import { makeTwoCellAnchor } from '../../src/drawing/anchor';
import { makeColor, makeSrgbColor } from '../../src/drawing/dml/colors';
import { makeSolidFill } from '../../src/drawing/dml/fill';
import { makePresetGeometry } from '../../src/drawing/dml/geometry';
import { makeShapeProperties } from '../../src/drawing/dml/shape-properties';
import { makeSimpleTextBody } from '../../src/drawing/dml/text';
import { makeChartDrawingItem, makeDrawing } from '../../src/drawing/drawing';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

const roundTrip = (d: ChartDrawing): ChartDrawing => parseUserShapesXml(userShapesToBytes(d));

describe('User-shapes XML round-trip', () => {
  it('preserves a relSizeAnchor with a text-box shape', () => {
    const d = makeChartDrawing([
      makeRelSizeAnchor(
        { x: 0.1, y: 0.2 },
        { x: 0.5, y: 0.6 },
        {
          kind: 'shape',
          shape: makeChartShape({
            id: 1,
            name: 'Annotation',
            descr: 'Top-left text box',
            txBox: true,
            spPr: makeShapeProperties({
              geometry: makePresetGeometry('rect'),
              fill: makeSolidFill(makeColor(makeSrgbColor('FFFFCC'))),
            }),
            txBody: makeSimpleTextBody('Hello, chart!', { sz: 1100, b: true }),
          }),
        },
      ),
    ]);
    const back = roundTrip(d);
    expect(back.shapes.length).toBe(1);
    const a = back.shapes[0];
    if (!a || a.kind !== 'relSize') throw new Error('expected relSize anchor');
    expect(a.from).toEqual({ x: 0.1, y: 0.2 });
    expect(a.to).toEqual({ x: 0.5, y: 0.6 });
    if (a.content.kind !== 'shape') throw new Error('expected shape content');
    expect(a.content.shape.name).toBe('Annotation');
    expect(a.content.shape.descr).toBe('Top-left text box');
    expect(a.content.shape.txBox).toBe(true);
    expect(a.content.shape.spPr?.fill).toEqual(
      makeSolidFill(makeColor(makeSrgbColor('FFFFCC'))),
    );
    expect(a.content.shape.txBody?.paragraphs[0]?.runs[0]).toMatchObject({
      kind: 'r',
      t: 'Hello, chart!',
    });
  });

  it('preserves an absSizeAnchor with EMU extent', () => {
    const d = makeChartDrawing([
      makeAbsSizeAnchor(
        { x: 0.05, y: 0.05 },
        { cx: 1828800, cy: 457200 },
        {
          kind: 'shape',
          shape: makeChartShape({ id: 2, name: 'Title', txBox: true }),
        },
      ),
    ]);
    const back = roundTrip(d);
    const a = back.shapes[0];
    if (!a || a.kind !== 'absSize') throw new Error('expected absSize anchor');
    expect(a.from).toEqual({ x: 0.05, y: 0.05 });
    expect(a.ext).toEqual({ cx: 1828800, cy: 457200 });
    if (a.content.kind !== 'shape') throw new Error('expected shape');
    expect(a.content.shape.id).toBe(2);
  });

  it('preserves a picture shape with embed rId', () => {
    const d = makeChartDrawing([
      makeRelSizeAnchor(
        { x: 0, y: 0 },
        { x: 0.3, y: 0.3 },
        {
          kind: 'picture',
          picture: { id: 3, name: 'Logo', embedRId: 'rId7' },
        },
      ),
    ]);
    const back = roundTrip(d);
    const a = back.shapes[0];
    if (!a || a.kind !== 'relSize') throw new Error('expected relSize');
    if (a.content.kind !== 'picture') throw new Error('expected picture');
    expect(a.content.picture.embedRId).toBe('rId7');
    expect(a.content.picture.name).toBe('Logo');
  });

  it('emits xmlns declarations on the root', () => {
    const xml = serializeUserShapes(makeChartDrawing());
    expect(xml).toContain('xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"');
    expect(xml).toContain('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"');
  });
});

describe('User-shapes workbook integration', () => {
  it('emits chartDrawing part + chart-rels + Override, restored on load', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } })] }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
    });
    space.userShapes = makeChartDrawing([
      makeRelSizeAnchor(
        { x: 0.1, y: 0.1 },
        { x: 0.4, y: 0.3 },
        {
          kind: 'shape',
          shape: makeChartShape({
            id: 1,
            name: 'Note',
            txBox: true,
            txBody: makeSimpleTextBody('Source: tests'),
          }),
        },
      ),
    ]);
    ws.drawing = makeDrawing([
      makeChartDrawingItem(makeTwoCellAnchor({ from: 'D2', to: 'J20' }), { space }),
    ]);

    const bytes = await workbookToBytes(wb);
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);

    expect(entries['xl/drawings/chartDrawing1.xml']).toBeDefined();
    expect(entries['xl/charts/_rels/chart1.xml.rels']).toBeDefined();

    const chartXml = new TextDecoder().decode(entries['xl/charts/chart1.xml']);
    expect(chartXml).toContain('<c:userShapes r:id="rId1"/>');

    const chartRels = new TextDecoder().decode(entries['xl/charts/_rels/chart1.xml.rels']);
    expect(chartRels).toContain('relationships/chartUserShapes');
    expect(chartRels).toContain('../drawings/chartDrawing1.xml');

    const ct = new TextDecoder().decode(entries['[Content_Types].xml']);
    expect(ct).toContain('/xl/drawings/chartDrawing1.xml');

    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const item = wb2.sheets[0]?.sheet.drawing?.items[0];
    if (!item || item.content.kind !== 'chart') throw new Error('expected chart item');
    const back = item.content.chart.space?.userShapes;
    expect(back?.shapes.length).toBe(1);
    const a = back?.shapes[0];
    if (!a || a.kind !== 'relSize') throw new Error('expected relSize anchor');
    if (a.content.kind !== 'shape') throw new Error('expected shape');
    expect(a.content.shape.name).toBe('Note');
    expect(a.content.shape.txBody?.paragraphs[0]?.runs[0]).toMatchObject({
      kind: 'r',
      t: 'Source: tests',
    });
  });
});
