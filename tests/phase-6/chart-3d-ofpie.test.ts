import { describe, expect, it } from 'vitest';
import {
  type Area3DChart,
  type Bar3DChart,
  type Line3DChart,
  makeArea3DChart,
  makeBar3DChart,
  makeBarSeries,
  makeChartSpace,
  makeLine3DChart,
  makeOfPieChart,
  makePie3DChart,
  makeSurface3DChart,
  type OfPieChart,
  type Pie3DChart,
  type Surface3DChart,
} from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';

const wrap = <T>(chart: { kind: string }, withAxes = true) =>
  makeChartSpace({
    plotArea: {
      // biome-ignore lint/suspicious/noExplicitAny: helper
      chart: chart as any,
      ...(withAxes ? { catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } } : {}),
    },
  }) as T;

describe('OfPieChart round-trip', () => {
  it('preserves ofPieType + secondPieSize + splitType + custSplit', () => {
    const chart = makeOfPieChart({
      ofPieType: 'bar',
      gapWidth: 100,
      secondPieSize: 80,
      splitType: 'cust',
      custSplit: [3, 5, 7],
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A10' } })],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart, false);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('ofPie');
    const ofPie = back.plotArea.chart as OfPieChart;
    expect(ofPie.ofPieType).toBe('bar');
    expect(ofPie.gapWidth).toBe(100);
    expect(ofPie.secondPieSize).toBe(80);
    expect(ofPie.splitType).toBe('cust');
    expect(ofPie.custSplit).toEqual([3, 5, 7]);
  });
});

describe('Bar3DChart round-trip', () => {
  it('preserves barDir + grouping + gapDepth + shape', () => {
    const chart = makeBar3DChart({
      barDir: 'col',
      grouping: 'stacked',
      gapWidth: 150,
      gapDepth: 200,
      shape: 'cylinder',
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } })],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('bar3D');
    const bar = back.plotArea.chart as Bar3DChart;
    expect(bar.barDir).toBe('col');
    expect(bar.grouping).toBe('stacked');
    expect(bar.gapDepth).toBe(200);
    expect(bar.shape).toBe('cylinder');
  });
});

describe('Line3DChart round-trip', () => {
  it('preserves grouping + gapDepth + per-series LineSeries.smooth', () => {
    const chart = makeLine3DChart({
      grouping: 'standard',
      gapDepth: 100,
      series: [
        makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }),
        { ...makeBarSeries({ idx: 1, val: { ref: 'B1:B4' } }), smooth: true },
      ],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('line3D');
    const line = back.plotArea.chart as Line3DChart;
    expect(line.gapDepth).toBe(100);
    expect(line.series[1]?.smooth).toBe(true);
  });
});

describe('Pie3DChart round-trip', () => {
  it('preserves single-series pie3D without axes', () => {
    const chart = makePie3DChart({
      varyColors: true,
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } })],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart, false);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('pie3D');
    const pie = back.plotArea.chart as Pie3DChart;
    expect(pie.varyColors).toBe(true);
    expect(pie.series.length).toBe(1);
  });
});

describe('Area3DChart round-trip', () => {
  it('preserves grouping + gapDepth', () => {
    const chart = makeArea3DChart({
      grouping: 'percentStacked',
      gapDepth: 50,
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } })],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('area3D');
    const area = back.plotArea.chart as Area3DChart;
    expect(area.grouping).toBe('percentStacked');
    expect(area.gapDepth).toBe(50);
  });
});

describe('Surface3DChart round-trip', () => {
  it('preserves wireframe + 3 axIds', () => {
    const chart = makeSurface3DChart({
      wireframe: false,
      axIds: [11, 22, 33],
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } })],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('surface3D');
    const surface = back.plotArea.chart as Surface3DChart;
    expect(surface.wireframe).toBe(false);
    expect(surface.axIds).toEqual([11, 22, 33]);
  });
});
