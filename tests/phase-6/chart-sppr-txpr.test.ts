import { describe, expect, it } from 'vitest';
import {
  type ChartSpace,
  makeBarChart,
  makeBarSeries,
  makeBubbleChart,
  makeBubbleSeries,
  makeChartSpace,
  makeScatterChart,
  makeScatterSeries,
} from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';
import { makeColor, makeSchemeColor, makeSrgbColor } from '../../src/drawing/dml/colors';
import { makeNoFill, makeSolidFill } from '../../src/drawing/dml/fill';
import { makePresetGeometry } from '../../src/drawing/dml/geometry';
import { makeLine } from '../../src/drawing/dml/line';
import { makeShapeProperties } from '../../src/drawing/dml/shape-properties';
import { makeParagraph, makeRun, makeSimpleTextBody, makeTextBody } from '../../src/drawing/dml/text';

const roundTrip = (s: ChartSpace): ChartSpace => parseChartXml(chartToBytes(s));

describe('Series spPr round-trip', () => {
  it('preserves BarSeries spPr (solid fill + outline)', () => {
    const sp = makeShapeProperties({
      fill: makeSolidFill(makeColor(makeSchemeColor('accent2'))),
      ln: makeLine({ w: 12700, fill: makeSolidFill(makeColor(makeSrgbColor('FF0000'))) }),
    });
    const back = roundTrip(
      makeChartSpace({
        plotArea: {
          chart: makeBarChart({
            series: [{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), spPr: sp }],
          }),
          catAx: { axId: 1, crossAx: 2 },
          valAx: { axId: 2, crossAx: 1 },
        },
      }),
    );
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.spPr).toEqual(sp);
  });

  it('preserves ScatterSeries spPr', () => {
    const sp = makeShapeProperties({ ln: makeLine({ w: 25400 }) });
    const space = makeChartSpace({
      plotArea: {
        chart: makeScatterChart({
          series: [
            {
              ...makeScatterSeries({ idx: 0, yVal: { ref: 'A1:A4' } }),
              spPr: sp,
            },
          ],
        }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
    });
    const back = roundTrip(space);
    if (back.plotArea.chart.kind !== 'scatter') throw new Error('expected scatter');
    expect(back.plotArea.chart.series[0]?.spPr).toEqual(sp);
  });

  it('preserves BubbleSeries spPr', () => {
    const sp = makeShapeProperties({ fill: makeNoFill() });
    const space = makeChartSpace({
      plotArea: {
        chart: makeBubbleChart({
          series: [
            {
              ...makeBubbleSeries({
                idx: 0,
                yVal: { ref: 'A1:A4' },
                bubbleSize: { ref: 'B1:B4' },
              }),
              spPr: sp,
            },
          ],
        }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
    });
    const back = roundTrip(space);
    if (back.plotArea.chart.kind !== 'bubble') throw new Error('expected bubble');
    expect(back.plotArea.chart.series[0]?.spPr).toEqual(sp);
  });
});

describe('Axis spPr / txPr round-trip', () => {
  it('preserves catAx + valAx spPr / txPr', () => {
    const axSpPr = makeShapeProperties({ ln: makeLine({ w: 6350 }) });
    const axTxPr = makeTextBody([makeParagraph([makeRun('', { sz: 900, b: false })])]);
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } })] }),
        catAx: { axId: 1, crossAx: 2, position: 'b', spPr: axSpPr, txPr: axTxPr },
        valAx: { axId: 2, crossAx: 1, position: 'l', spPr: axSpPr, txPr: axTxPr },
      },
    });
    const back = roundTrip(space);
    expect(back.plotArea.catAx?.spPr).toEqual(axSpPr);
    expect(back.plotArea.catAx?.txPr).toEqual(axTxPr);
    expect(back.plotArea.valAx?.spPr).toEqual(axSpPr);
    expect(back.plotArea.valAx?.txPr).toEqual(axTxPr);
  });
});

describe('Legend spPr / txPr round-trip', () => {
  it('preserves legend overlay + spPr + txPr', () => {
    const legendSpPr = makeShapeProperties({
      fill: makeSolidFill(makeColor(makeSrgbColor('F0F0F0'))),
    });
    const legendTxPr = makeTextBody([makeParagraph([makeRun('', { sz: 800 })])]);
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A2' } })] }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
      legend: { position: 'b', overlay: true, spPr: legendSpPr, txPr: legendTxPr },
    });
    const back = roundTrip(space);
    expect(back.legend?.position).toBe('b');
    expect(back.legend?.overlay).toBe(true);
    expect(back.legend?.spPr).toEqual(legendSpPr);
    expect(back.legend?.txPr).toEqual(legendTxPr);
  });
});

describe('ChartTitle round-trip', () => {
  it('preserves plain string title (legacy shape)', () => {
    const back = roundTrip(
      makeChartSpace({
        plotArea: {
          chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A2' } })] }),
          catAx: { axId: 1, crossAx: 2 },
          valAx: { axId: 2, crossAx: 1 },
        },
        title: 'Hello',
      }),
    );
    expect(back.title?.text).toBe('Hello');
  });

  it('preserves ChartTitle.tx (rich body) + spPr + overlay', () => {
    const tx = makeSimpleTextBody('Rich Title', { sz: 1800, b: true });
    const sp = makeShapeProperties({
      ln: makeLine({ w: 12700, fill: makeSolidFill(makeColor(makeSrgbColor('123456'))) }),
    });
    const back = roundTrip(
      makeChartSpace({
        plotArea: {
          chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A2' } })] }),
          catAx: { axId: 1, crossAx: 2 },
          valAx: { axId: 2, crossAx: 1 },
        },
        title: { tx, overlay: true, spPr: sp },
      }),
    );
    expect(back.title?.text).toBe('Rich Title');
    expect(back.title?.overlay).toBe(true);
    expect(back.title?.spPr).toEqual(sp);
    // The parser also fills `tx` so callers get the full body for re-edits.
    expect(back.title?.tx?.paragraphs[0]?.runs[0]).toMatchObject({ kind: 'r', t: 'Rich Title' });
  });
});

describe('PlotArea spPr round-trip', () => {
  it('preserves plot-area background fill', () => {
    const sp = makeShapeProperties({
      fill: makeSolidFill(makeColor(makeSrgbColor('FAFAFA'))),
      geometry: makePresetGeometry('rect'),
    });
    const back = roundTrip(
      makeChartSpace({
        plotArea: {
          chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A2' } })] }),
          catAx: { axId: 1, crossAx: 2 },
          valAx: { axId: 2, crossAx: 1 },
          spPr: sp,
        },
      }),
    );
    expect(back.plotArea.spPr).toEqual(sp);
  });
});

describe('ChartSpace top-level spPr / txPr round-trip', () => {
  it('preserves chartSpace-level outer frame + default text', () => {
    const sp = makeShapeProperties({ ln: makeLine({ w: 6350 }) });
    const txPr = makeTextBody([makeParagraph([makeRun('', { sz: 1000 })])]);
    const back = roundTrip(
      makeChartSpace({
        plotArea: {
          chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A2' } })] }),
          catAx: { axId: 1, crossAx: 2 },
          valAx: { axId: 2, crossAx: 1 },
        },
        spPr: sp,
        txPr,
      }),
    );
    expect(back.spPr).toEqual(sp);
    expect(back.txPr).toEqual(txPr);
  });
});
