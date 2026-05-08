import { describe, expect, it } from 'vitest';
import {
  type ChartSpace,
  type DataLabelList,
  type ErrorBars,
  makeBarChart,
  makeBarSeries,
  makeBubbleChart,
  makeBubbleSeries,
  makeChartSpace,
  makeScatterChart,
  makeScatterSeries,
  type Trendline,
} from '../../src/xlsx/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/xlsx/chart/chart-xml';
import { makeColor, makeSrgbColor } from '../../src/xlsx/drawing/dml/colors';
import { makeSolidFill } from '../../src/xlsx/drawing/dml/fill';
import { makeLine } from '../../src/xlsx/drawing/dml/line';
import { makeShapeProperties } from '../../src/xlsx/drawing/dml/shape-properties';
import { makeParagraph, makeRun, makeSimpleTextBody, makeTextBody } from '../../src/xlsx/drawing/dml/text';

const roundTrip = (s: ChartSpace): ChartSpace => parseChartXml(chartToBytes(s));

const wrapBarSeries = (series: NonNullable<Parameters<typeof makeBarChart>[0]['series']>): ChartSpace =>
  makeChartSpace({
    plotArea: {
      chart: makeBarChart({ series }),
      catAx: { axId: 1, crossAx: 2 },
      valAx: { axId: 2, crossAx: 1 },
    },
  });

describe('DataLabelList round-trip', () => {
  it('preserves series-wide showVal / dLblPos / numFmt', () => {
    const dLbls: DataLabelList = {
      showVal: true,
      showCatName: false,
      showSerName: true,
      showPercent: false,
      showLegendKey: false,
      showBubbleSize: false,
      dLblPos: 'outEnd',
      separator: ', ',
      numFmt: { formatCode: '0.00%', sourceLinked: false },
      showLeaderLines: true,
    };
    const back = roundTrip(wrapBarSeries([{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), dLbls }]));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.dLbls).toEqual(dLbls);
  });

  it('preserves delete=1 short-circuit (no other children)', () => {
    const back = roundTrip(
      wrapBarSeries([
        {
          ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }),
          dLbls: { delete: true },
        },
      ]),
    );
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.dLbls).toEqual({ delete: true });
  });

  it('preserves per-point dLbl with rich tx + dLblPos override', () => {
    const dLbls: DataLabelList = {
      showVal: true,
      dLbl: [
        {
          idx: 2,
          tx: { kind: 'rich', body: makeSimpleTextBody('Highlight', { sz: 1200, b: true }) },
          dLblPos: 'inEnd',
          showVal: true,
        },
      ],
    };
    const back = roundTrip(wrapBarSeries([{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), dLbls }]));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    const back0 = back.plotArea.chart.series[0]?.dLbls;
    expect(back0?.dLbl?.length).toBe(1);
    expect(back0?.dLbl?.[0]?.idx).toBe(2);
    expect(back0?.dLbl?.[0]?.dLblPos).toBe('inEnd');
    if (back0?.dLbl?.[0]?.tx?.kind !== 'rich') throw new Error('expected rich tx');
    expect(back0.dLbl[0].tx.body.paragraphs[0]?.runs[0]).toMatchObject({ kind: 'r', t: 'Highlight' });
  });

  it('preserves per-point dLbl with strRef tx', () => {
    const dLbls: DataLabelList = {
      dLbl: [{ idx: 0, tx: { kind: 'strRef', ref: 'Sheet1!$A$5' } }],
    };
    const back = roundTrip(wrapBarSeries([{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), dLbls }]));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.dLbls?.dLbl?.[0]?.tx).toEqual({
      kind: 'strRef',
      ref: 'Sheet1!$A$5',
    });
  });

  it('preserves dLbl spPr / txPr formatting', () => {
    const dLbls: DataLabelList = {
      spPr: makeShapeProperties({ fill: makeSolidFill(makeColor(makeSrgbColor('FFFFEE'))) }),
      txPr: makeTextBody([makeParagraph([makeRun('', { sz: 900 })])]),
    };
    const back = roundTrip(wrapBarSeries([{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), dLbls }]));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.dLbls?.spPr).toEqual(dLbls.spPr);
    expect(back.plotArea.chart.series[0]?.dLbls?.txPr).toEqual(dLbls.txPr);
  });
});

describe('Trendline round-trip', () => {
  it('preserves linear trendline with name / forecast / dispEq / dispRSqr', () => {
    const t: Trendline = {
      name: 'Linear (Sales)',
      trendlineType: 'linear',
      forward: 2,
      backward: 0.5,
      intercept: 10,
      dispEq: true,
      dispRSqr: true,
    };
    const back = roundTrip(wrapBarSeries([{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), trendline: [t] }]));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.trendline).toEqual([t]);
  });

  it('preserves polynomial trendline with order', () => {
    const t: Trendline = { trendlineType: 'poly', order: 3 };
    const back = roundTrip(wrapBarSeries([{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), trendline: [t] }]));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.trendline?.[0]?.order).toBe(3);
  });

  it('preserves moving-average trendline with period', () => {
    const t: Trendline = { trendlineType: 'movingAvg', period: 5 };
    const back = roundTrip(wrapBarSeries([{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), trendline: [t] }]));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.trendline?.[0]).toEqual({
      trendlineType: 'movingAvg',
      period: 5,
    });
  });

  it('preserves multiple trendlines on one series + their spPr', () => {
    const trends: Trendline[] = [
      { trendlineType: 'linear', name: 'Lin', dispEq: true },
      {
        trendlineType: 'exp',
        name: 'Exp',
        spPr: makeShapeProperties({ ln: makeLine({ w: 12700 }) }),
      },
    ];
    const back = roundTrip(wrapBarSeries([{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), trendline: trends }]));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.trendline?.length).toBe(2);
    expect(back.plotArea.chart.series[0]?.trendline?.[0]?.trendlineType).toBe('linear');
    expect(back.plotArea.chart.series[0]?.trendline?.[1]?.spPr?.ln?.w).toBe(12700);
  });
});

describe('ErrorBars round-trip', () => {
  it('preserves bar-series y-direction errBars (fixedVal)', () => {
    const e: ErrorBars = {
      errDir: 'y',
      errBarType: 'both',
      errValType: 'fixedVal',
      val: 5,
      noEndCap: false,
    };
    const back = roundTrip(wrapBarSeries([{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), errBars: [e] }]));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.errBars).toEqual([e]);
  });

  it('preserves custom errBars with plus / minus NumericRef', () => {
    const e: ErrorBars = {
      errBarType: 'both',
      errValType: 'cust',
      plus: { ref: 'Sheet1!$B$1:$B$4', cache: [1, 1, 1, 1] },
      minus: { ref: 'Sheet1!$C$1:$C$4', cache: [0.5, 0.5, 0.5, 0.5] },
    };
    const back = roundTrip(wrapBarSeries([{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), errBars: [e] }]));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.errBars?.[0]?.plus?.ref).toBe('Sheet1!$B$1:$B$4');
    expect(back.plotArea.chart.series[0]?.errBars?.[0]?.minus?.cache).toEqual([0.5, 0.5, 0.5, 0.5]);
  });

  it('preserves scatter x + y errBars (two entries in document order)', () => {
    const space = makeChartSpace({
      plotArea: {
        chart: makeScatterChart({
          series: [
            {
              ...makeScatterSeries({ idx: 0, yVal: { ref: 'A1:A4' } }),
              errBars: [
                { errDir: 'x', errBarType: 'both', errValType: 'percentage', val: 10 },
                { errDir: 'y', errBarType: 'both', errValType: 'stdErr' },
              ],
            },
          ],
        }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
    });
    const back = roundTrip(space);
    if (back.plotArea.chart.kind !== 'scatter') throw new Error('expected scatter');
    const ebs = back.plotArea.chart.series[0]?.errBars;
    expect(ebs?.length).toBe(2);
    expect(ebs?.[0]?.errDir).toBe('x');
    expect(ebs?.[0]?.errValType).toBe('percentage');
    expect(ebs?.[0]?.val).toBe(10);
    expect(ebs?.[1]?.errDir).toBe('y');
    expect(ebs?.[1]?.errValType).toBe('stdErr');
  });
});

describe('Bubble series carries decorations', () => {
  it('preserves bubble series dLbls / trendline / errBars', () => {
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
              dLbls: { showBubbleSize: true, showVal: false },
              trendline: [{ trendlineType: 'linear' }],
              errBars: [{ errDir: 'y', errBarType: 'plus', errValType: 'stdDev' }],
            },
          ],
        }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
    });
    const back = roundTrip(space);
    if (back.plotArea.chart.kind !== 'bubble') throw new Error('expected bubble');
    const s = back.plotArea.chart.series[0];
    expect(s?.dLbls?.showBubbleSize).toBe(true);
    expect(s?.dLbls?.showVal).toBe(false);
    expect(s?.trendline?.[0]?.trendlineType).toBe('linear');
    expect(s?.errBars?.[0]?.errBarType).toBe('plus');
  });
});

describe('Document ordering (ECMA-376)', () => {
  it('series children appear in tx → spPr → dLbls → trendline → errBars → cat → val order', () => {
    const space = wrapBarSeries([
      {
        ...makeBarSeries({
          idx: 0,
          val: { ref: 'A1:A4' },
          cat: { ref: 'B1:B4', cacheKind: 'str' },
          tx: { kind: 'literal', value: 'Sales' },
        }),
        spPr: makeShapeProperties({
          fill: makeSolidFill(makeColor(makeSrgbColor('FF0000'))),
        }),
        dLbls: { showVal: true },
        trendline: [{ trendlineType: 'linear' }],
        errBars: [{ errBarType: 'both', errValType: 'fixedVal', val: 1 }],
      },
    ]);
    const xml = new TextDecoder().decode(chartToBytes(space));
    const idxes = ['<c:tx>', '<c:spPr>', '<c:dLbls>', '<c:trendline>', '<c:errBars>', '<c:cat>', '<c:val>'].map((t) =>
      xml.indexOf(t),
    );
    for (let i = 0; i < idxes.length; i++) expect(idxes[i]).toBeGreaterThan(-1);
    for (let i = 1; i < idxes.length; i++) {
      expect(idxes[i]).toBeGreaterThan(idxes[i - 1] as number);
    }
  });
});
