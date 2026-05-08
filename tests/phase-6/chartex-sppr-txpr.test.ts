import { describe, expect, it } from 'vitest';
import {
  type CxChartSpace,
  makeCxChartSpace,
  makeCxData,
  makeCxNumDim,
  makeCxSeries,
  makeCxStrDim,
  makeWaterfallChart,
} from '../../src/xlsx/chart/cx/chartex';
import { chartExToBytes, parseChartExXml, serializeChartExSpace } from '../../src/xlsx/chart/cx/chartex-xml';
import { makeColor, makeSchemeColor, makeSrgbColor } from '../../src/xlsx/drawing/dml/colors';
import { makeSolidFill } from '../../src/xlsx/drawing/dml/fill';
import { makeLine } from '../../src/xlsx/drawing/dml/line';
import { makeShapeProperties } from '../../src/xlsx/drawing/dml/shape-properties';
import { makeParagraph, makeRun, makeTextBody } from '../../src/xlsx/drawing/dml/text';

const roundTrip = (s: CxChartSpace): CxChartSpace => parseChartExXml(chartExToBytes(s));

const wrap = (overrides: Partial<Parameters<typeof makeCxChartSpace>[0]> = {}): CxChartSpace =>
  makeCxChartSpace({
    data: [
      makeCxData(0, [makeCxStrDim({ type: 'cat', f: 'A1:A3' }), makeCxNumDim({ type: 'val', f: 'B1:B3' })]),
    ],
    series: [makeCxSeries({ layoutId: 'sunburst', dataId: 0 })],
    ...overrides,
  });

describe('chartex Series spPr / txPr', () => {
  it('preserves series spPr (solid fill) round-trip', () => {
    const sp = makeShapeProperties({
      fill: makeSolidFill(makeColor(makeSchemeColor('accent2'))),
      ln: makeLine({ w: 12700 }),
    });
    const back = roundTrip(
      wrap({
        series: [
          makeCxSeries({
            layoutId: 'sunburst',
            dataId: 0,
            spPr: sp,
          }),
        ],
      }),
    );
    expect(back.chart.plotArea.series[0]?.spPr).toEqual(sp);
  });

  it('preserves series txPr round-trip', () => {
    const txPr = makeTextBody([makeParagraph([makeRun('', { sz: 1000 })])]);
    const back = roundTrip(
      wrap({
        series: [
          makeCxSeries({ layoutId: 'sunburst', dataId: 0, txPr }),
        ],
      }),
    );
    expect(back.chart.plotArea.series[0]?.txPr).toEqual(txPr);
  });
});

describe('chartex Axis spPr / txPr', () => {
  it('preserves axis spPr + txPr on both axes of a Waterfall chart', () => {
    const axSpPr = makeShapeProperties({ ln: makeLine({ w: 6350 }) });
    const axTxPr = makeTextBody([makeParagraph([makeRun('', { sz: 900 })])]);
    const space = makeWaterfallChart({ catRef: 'A1:A4', valRef: 'B1:B4' });
    space.chart.plotArea.axes = space.chart.plotArea.axes.map((a) => ({
      ...a,
      spPr: axSpPr,
      txPr: axTxPr,
    }));
    const back = roundTrip(space);
    for (const a of back.chart.plotArea.axes) {
      expect(a.spPr).toEqual(axSpPr);
      expect(a.txPr).toEqual(axTxPr);
    }
  });
});

describe('chartex Title spPr / txPr', () => {
  it('preserves title spPr + txPr alongside text', () => {
    const sp = makeShapeProperties({ fill: makeSolidFill(makeColor(makeSrgbColor('FAFAFA'))) });
    const txPr = makeTextBody([makeParagraph([makeRun('', { sz: 1800, b: true })])]);
    const back = roundTrip(
      wrap({
        title: { pos: 't', align: 'ctr', overlay: false, text: 'My Chart', spPr: sp, txPr },
      }),
    );
    expect(back.chart.title?.text).toBe('My Chart');
    expect(back.chart.title?.spPr).toEqual(sp);
    expect(back.chart.title?.txPr).toEqual(txPr);
  });
});

describe('chartex Legend spPr / txPr', () => {
  it('preserves legend spPr + txPr', () => {
    const sp = makeShapeProperties({ fill: makeSolidFill(makeColor(makeSrgbColor('F0F0F0'))) });
    const txPr = makeTextBody([makeParagraph([makeRun('', { sz: 800 })])]);
    const back = roundTrip(
      wrap({
        legend: { pos: 'b', align: 'ctr', overlay: false, spPr: sp, txPr },
      }),
    );
    expect(back.chart.legend?.pos).toBe('b');
    expect(back.chart.legend?.spPr).toEqual(sp);
    expect(back.chart.legend?.txPr).toEqual(txPr);
  });
});

describe('chartex PlotArea (plotSurface) spPr', () => {
  it('preserves plot-surface background spPr', () => {
    const sp = makeShapeProperties({ fill: makeSolidFill(makeColor(makeSrgbColor('FFFFEE'))) });
    const back = roundTrip(wrap({ plotAreaSpPr: sp }));
    expect(back.chart.plotArea.spPr).toEqual(sp);
  });

  it('emits plotSurface as the first child of plotAreaRegion', () => {
    const sp = makeShapeProperties({ ln: makeLine({ w: 9525 }) });
    const xml = serializeChartExSpace(wrap({ plotAreaSpPr: sp }));
    const region = xml.indexOf('<cx:plotAreaRegion>');
    const plotSurface = xml.indexOf('<cx:plotSurface');
    const series = xml.indexOf('<cx:series ');
    expect(region).toBeGreaterThan(-1);
    expect(plotSurface).toBeGreaterThan(region);
    expect(series).toBeGreaterThan(plotSurface);
  });
});

describe('chartex ChartSpace top-level spPr / txPr', () => {
  it('preserves chartSpace-level spPr + txPr', () => {
    const sp = makeShapeProperties({ ln: makeLine({ w: 6350 }) });
    const txPr = makeTextBody([makeParagraph([makeRun('', { sz: 1000 })])]);
    const back = roundTrip(wrap({ spPr: sp, txPr }));
    expect(back.spPr).toEqual(sp);
    expect(back.txPr).toEqual(txPr);
  });
});
