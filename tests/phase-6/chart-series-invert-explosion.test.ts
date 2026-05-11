import { describe, expect, it } from 'vitest';
import {
  type BarChart,
  type BubbleChart,
  type PieChart,
  makeBarChart,
  makeBarSeries,
  makeBubbleChart,
  makeBubbleSeries,
  makeChartSpace,
  makePieChart,
} from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

describe('series-level invertIfNegative + explosion', () => {
  it('serialises BarSeries.invertIfNegative after <c:spPr>', () => {
    const chart = makeBarChart({
      series: [
        { ...makeBarSeries({ idx: 0, val: { ref: 'A1:A3' } }), invertIfNegative: true },
      ],
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).toContain('<c:invertIfNegative val="1"/>');
  });

  it('round-trips BarSeries.invertIfNegative', () => {
    const chart = makeBarChart({
      series: [
        { ...makeBarSeries({ idx: 0, val: { ref: 'A1:A3' } }), invertIfNegative: true },
      ],
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const back = parseChartXml(chartToBytes(space));
    const bar = back.plotArea.chart as BarChart;
    expect(bar.series[0]?.invertIfNegative).toBe(true);
  });

  it('round-trips PieSeries.explosion', () => {
    const chart = makePieChart({
      series: [{ ...makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } }), explosion: 30 }],
    });
    const space = makeChartSpace({ plotArea: { chart } });
    const back = parseChartXml(chartToBytes(space));
    const pie = back.plotArea.chart as PieChart;
    expect(pie.series[0]?.explosion).toBe(30);
  });

  it('round-trips BubbleSeries.invertIfNegative', () => {
    const chart = makeBubbleChart({
      series: [
        {
          ...makeBubbleSeries({
            idx: 0,
            yVal: { ref: 'B1:B3' },
            bubbleSize: { ref: 'C1:C3' },
          }),
          invertIfNegative: true,
        },
      ],
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const back = parseChartXml(chartToBytes(space));
    const bubble = back.plotArea.chart as BubbleChart;
    expect(bubble.series[0]?.invertIfNegative).toBe(true);
  });

  it('omits both fields when unset', () => {
    const chart = makeBarChart({
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A3' } })],
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).not.toContain('<c:invertIfNegative');
    expect(xml).not.toContain('<c:explosion');
  });
});
