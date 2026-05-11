import { describe, expect, it } from 'vitest';
import {
  type LineChart,
  type LineSeries,
  type ScatterChart,
  type ScatterSeries,
  makeBarSeries,
  makeChartSpace,
  makeLineChart,
  makeScatterChart,
  makeScatterSeries,
} from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';
import { makeColor, makeSrgbColor } from '../../src/drawing/dml/colors';
import { makeSolidFill } from '../../src/drawing/dml/fill';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

describe('issue #47 — series.marker is exposed and round-trips', () => {
  it('serialises <c:marker> inside a LineSeries with symbol/size/spPr', () => {
    const series: LineSeries = {
      ...makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } }),
      marker: {
        symbol: 'circle',
        size: 7,
        spPr: { fill: makeSolidFill(makeColor(makeSrgbColor('4263EB'))) },
      },
    };
    const chart = makeLineChart({ series: [series] });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).toContain('<c:marker>');
    expect(xml).toContain('<c:symbol val="circle"/>');
    expect(xml).toContain('<c:size val="7"/>');
    expect(xml).toContain('<a:srgbClr val="4263EB"');
    // <c:marker> must sit between <c:spPr> and (any subsequent) <c:dLbls>/<c:val>
    const markerPos = xml.indexOf('<c:marker>');
    const valPos = xml.indexOf('<c:val>');
    expect(markerPos).toBeGreaterThan(0);
    expect(markerPos).toBeLessThan(valPos);
  });

  it('round-trips a LineSeries marker through parseChartXml', () => {
    const series: LineSeries = {
      ...makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } }),
      marker: { symbol: 'diamond', size: 5 },
    };
    const chart = makeLineChart({ series: [series] });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const back = parseChartXml(chartToBytes(space));
    const line = back.plotArea.chart as LineChart;
    expect(line.series[0]?.marker).toEqual({ symbol: 'diamond', size: 5 });
  });

  it('round-trips a ScatterSeries marker through parseChartXml', () => {
    const series: ScatterSeries = makeScatterSeries({
      idx: 0,
      yVal: { ref: 'Sheet1!$B$2:$B$5' },
      xVal: { ref: 'Sheet1!$A$2:$A$5' },
      marker: { symbol: 'triangle', size: 9 },
    });
    const chart = makeScatterChart({ series: [series] });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const back = parseChartXml(chartToBytes(space));
    const scatter = back.plotArea.chart as ScatterChart;
    expect(scatter.series[0]?.marker).toEqual({ symbol: 'triangle', size: 9 });
  });

  it('omits <c:marker> when unset', () => {
    const chart = makeLineChart({
      series: [{ ...makeBarSeries({ idx: 0, val: { ref: 'X1:X1' } }) }],
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).not.toContain('<c:marker>');
  });
});
