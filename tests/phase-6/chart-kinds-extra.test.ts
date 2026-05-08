import { describe, expect, it } from 'vitest';
import {
  type BubbleChart,
  makeBarSeries,
  makeBubbleChart,
  makeBubbleSeries,
  makeChartSpace,
  makeStockChart,
  makeSurfaceChart,
  type StockChart,
  type SurfaceChart,
} from '../../src/xlsx/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/xlsx/chart/chart-xml';

const wrap = <T>(chart: { kind: string }, withAxes = true) =>
  makeChartSpace({
    plotArea: {
      // biome-ignore lint/suspicious/noExplicitAny: helper for varied chart kinds
      chart: chart as any,
      ...(withAxes ? { catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } } : {}),
    },
  }) as T;

describe('BubbleChart round-trip', () => {
  it('preserves bubble3D / bubbleScale / sizeRepresents + xVal/yVal/bubbleSize', () => {
    const chart = makeBubbleChart({
      bubble3D: true,
      bubbleScale: 75,
      showNegBubbles: false,
      sizeRepresents: 'area',
      series: [
        makeBubbleSeries({
          idx: 0,
          xVal: { ref: 'A1:A3', cache: [1, 2, 3] },
          yVal: { ref: 'B1:B3', cache: [10, 20, 30] },
          bubbleSize: { ref: 'C1:C3', cache: [5, 10, 15] },
          bubble3D: true,
        }),
      ],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('bubble');
    const bubble = back.plotArea.chart as BubbleChart;
    expect(bubble.bubble3D).toBe(true);
    expect(bubble.bubbleScale).toBe(75);
    expect(bubble.showNegBubbles).toBe(false);
    expect(bubble.sizeRepresents).toBe('area');
    const s = bubble.series[0];
    expect(s?.xVal?.cache).toEqual([1, 2, 3]);
    expect(s?.yVal.cache).toEqual([10, 20, 30]);
    expect(s?.bubbleSize.cache).toEqual([5, 10, 15]);
    expect(s?.bubble3D).toBe(true);
  });
});

describe('StockChart round-trip', () => {
  it('preserves hiLowLines + upDownBars presence + 4 series', () => {
    const chart = makeStockChart({
      hiLowLines: true,
      upDownBars: true,
      series: [
        makeBarSeries({ idx: 0, val: { ref: 'B2:B6' } }),
        makeBarSeries({ idx: 1, val: { ref: 'C2:C6' } }),
        makeBarSeries({ idx: 2, val: { ref: 'D2:D6' } }),
        makeBarSeries({ idx: 3, val: { ref: 'E2:E6' } }),
      ],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('stock');
    const stock = back.plotArea.chart as StockChart;
    expect(stock.series.length).toBe(4);
    expect(stock.hiLowLines).toBe(true);
    expect(stock.upDownBars).toBe(true);
  });

  it('omits hiLowLines / upDownBars when not set', () => {
    const chart = makeStockChart({
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A3' } })],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    const stock = back.plotArea.chart as StockChart;
    expect(stock.hiLowLines).toBeUndefined();
    expect(stock.upDownBars).toBeUndefined();
  });
});

describe('SurfaceChart round-trip', () => {
  it('preserves wireframe + 3 axIds', () => {
    const chart = makeSurfaceChart({
      wireframe: true,
      axIds: [10, 20, 30],
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } })],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('surface');
    const surface = back.plotArea.chart as SurfaceChart;
    expect(surface.wireframe).toBe(true);
    expect(surface.axIds).toEqual([10, 20, 30]);
  });
});
