import { describe, expect, it } from 'vitest';
import {
  type BarChart,
  type BubbleChart,
  type DoughnutChart,
  type PieChart,
  type ScatterChart,
  makeBarChart,
  makeBarSeries,
  makeBubbleChart,
  makeBubbleSeries,
  makeChartSpace,
  makeDoughnutChart,
  makePieChart,
  makeScatterChart,
  makeScatterSeries,
} from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';
import { makeColor, makeSrgbColor } from '../../src/drawing/dml/colors';
import { makeSolidFill } from '../../src/drawing/dml/fill';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

const redFill = () => ({ fill: makeSolidFill(makeColor(makeSrgbColor('FF0000'))) });
const blueFill = () => ({ fill: makeSolidFill(makeColor(makeSrgbColor('0000FF'))) });

describe('issue #44 — per-point <c:dPt> data points round-trip', () => {
  it('serialises <c:dPt> children inside a BarSeries', () => {
    const chart = makeBarChart({
      series: [
        {
          ...makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$4' } }),
          dPt: [
            { idx: 0, spPr: redFill() },
            { idx: 2, spPr: blueFill(), invertIfNegative: true },
          ],
        },
      ],
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).toContain('<c:dPt><c:idx val="0"/>');
    expect(xml).toContain('<c:dPt><c:idx val="2"/><c:invertIfNegative val="1"/>');
    // dPt must sit before val per ECMA-376 sequence
    const dptPos = xml.indexOf('<c:dPt>');
    const valPos = xml.indexOf('<c:val>');
    expect(dptPos).toBeGreaterThan(0);
    expect(dptPos).toBeLessThan(valPos);
  });

  it('round-trips dPt through parseChartXml on a Pie chart', () => {
    const chart = makePieChart({
      series: [
        {
          ...makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$4' } }),
          dPt: [
            { idx: 0, spPr: redFill(), explosion: 25 },
            { idx: 1, spPr: blueFill() },
          ],
        },
      ],
    });
    const space = makeChartSpace({ plotArea: { chart } });
    const back = parseChartXml(chartToBytes(space));
    const pie = back.plotArea.chart as PieChart;
    const dPt = pie.series[0]?.dPt;
    expect(dPt?.length).toBe(2);
    expect(dPt?.[0]?.idx).toBe(0);
    expect(dPt?.[0]?.explosion).toBe(25);
    expect(dPt?.[1]?.idx).toBe(1);
  });

  it('round-trips dPt on a Doughnut chart', () => {
    const chart = makeDoughnutChart({
      series: [
        {
          ...makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } }),
          dPt: [{ idx: 0, spPr: redFill() }],
        },
      ],
    });
    const space = makeChartSpace({ plotArea: { chart } });
    const back = parseChartXml(chartToBytes(space));
    const donut = back.plotArea.chart as DoughnutChart;
    expect(donut.series[0]?.dPt?.[0]?.idx).toBe(0);
  });

  it('round-trips dPt on a Scatter chart with marker overrides', () => {
    const chart = makeScatterChart({
      series: [
        {
          ...makeScatterSeries({ idx: 0, yVal: { ref: 'Sheet1!$B$2:$B$5' } }),
          dPt: [{ idx: 1, marker: { symbol: 'star', size: 10 } }],
        },
      ],
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const back = parseChartXml(chartToBytes(space));
    const scatter = back.plotArea.chart as ScatterChart;
    expect(scatter.series[0]?.dPt?.[0]?.marker).toEqual({ symbol: 'star', size: 10 });
  });

  it('round-trips bubble3D on a BubbleSeries dPt', () => {
    const chart = makeBubbleChart({
      series: [
        {
          ...makeBubbleSeries({
            idx: 0,
            yVal: { ref: 'Sheet1!$B$2:$B$5' },
            bubbleSize: { ref: 'Sheet1!$C$2:$C$5' },
          }),
          dPt: [{ idx: 0, bubble3D: true, spPr: redFill() }],
        },
      ],
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const back = parseChartXml(chartToBytes(space));
    const bubble = back.plotArea.chart as BubbleChart;
    expect(bubble.series[0]?.dPt?.[0]?.bubble3D).toBe(true);
  });

  it('omits <c:dPt> when none are set', () => {
    const chart = makeBarChart({
      series: [makeBarSeries({ idx: 0, val: { ref: 'X1:X1' } })],
    });
    const space = makeChartSpace({
      plotArea: { chart, catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).not.toContain('<c:dPt>');
    // existing chart parses still work
    const back = parseChartXml(chartToBytes(space));
    const bar = back.plotArea.chart as BarChart;
    expect(bar.series[0]?.dPt).toBeUndefined();
  });
});
