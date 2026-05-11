import { describe, expect, it } from 'vitest';
import { type BarChart, makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

describe('issue #45 — BarChart.overlap is exposed and round-trips', () => {
  it('serialises an explicit overlap value', () => {
    const chart = makeBarChart({
      barDir: 'col',
      grouping: 'stacked',
      series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } })],
      overlap: 100,
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).toContain('<c:overlap val="100"/>');
  });

  it('round-trips overlap through parseChartXml', () => {
    const chart = makeBarChart({
      grouping: 'clustered',
      series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } })],
      overlap: -25,
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const back = parseChartXml(chartToBytes(space));
    const bar = back.plotArea.chart as BarChart;
    expect(bar.overlap).toBe(-25);
  });

  it('falls back to the stacked default of 100 when overlap is unset', () => {
    const chart = makeBarChart({
      grouping: 'percentStacked',
      series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } })],
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).toContain('<c:overlap val="100"/>');
  });

  it('omits <c:overlap> for clustered when overlap is unset', () => {
    const chart = makeBarChart({
      grouping: 'clustered',
      series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } })],
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).not.toContain('<c:overlap');
  });
});
