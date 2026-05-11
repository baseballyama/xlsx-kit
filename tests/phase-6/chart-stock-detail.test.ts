import { describe, expect, it } from 'vitest';
import {
  type StockChart,
  makeBarSeries,
  makeChartSpace,
  makeStockChart,
} from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';
import { makeColor, makeSrgbColor } from '../../src/drawing/dml/colors';
import { makeSolidFill } from '../../src/drawing/dml/fill';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

const stockSpace = (
  hiLowLines: StockChart['hiLowLines'],
  upDownBars: StockChart['upDownBars'],
) =>
  makeChartSpace({
    plotArea: {
      chart: makeStockChart({
        series: [
          makeBarSeries({ idx: 0, val: { ref: 'A1:A5' } }),
          makeBarSeries({ idx: 1, val: { ref: 'B1:B5' } }),
          makeBarSeries({ idx: 2, val: { ref: 'C1:C5' } }),
          makeBarSeries({ idx: 3, val: { ref: 'D1:D5' } }),
        ],
        ...(hiLowLines !== undefined ? { hiLowLines } : {}),
        ...(upDownBars !== undefined ? { upDownBars } : {}),
      }),
      catAx: { axId: 1, crossAx: 2 },
      valAx: { axId: 2, crossAx: 1 },
    },
  });

describe('StockChart hiLowLines + upDownBars detail', () => {
  it('keeps emitting <c:hiLowLines/> for the boolean form (backward compat)', () => {
    const xml = decode(chartToBytes(stockSpace(true, undefined)));
    expect(xml).toContain('<c:hiLowLines/>');
  });

  it('emits <c:hiLowLines><c:spPr>...</c:hiLowLines> for the detailed form', () => {
    const xml = decode(
      chartToBytes(
        stockSpace({ spPr: { fill: makeSolidFill(makeColor(makeSrgbColor('888888'))) } }, undefined),
      ),
    );
    expect(xml).toContain('<c:hiLowLines><c:spPr>');
    expect(xml).toContain('<a:srgbClr val="888888"');
  });

  it('emits <c:upDownBars> with gapWidth + upBars/downBars detail', () => {
    const xml = decode(
      chartToBytes(
        stockSpace(undefined, {
          gapWidth: 100,
          upBars: { spPr: { fill: makeSolidFill(makeColor(makeSrgbColor('22AA22'))) } },
          downBars: { spPr: { fill: makeSolidFill(makeColor(makeSrgbColor('AA2222'))) } },
        }),
      ),
    );
    expect(xml).toContain('<c:upDownBars>');
    expect(xml).toContain('<c:gapWidth val="100"/>');
    expect(xml).toContain('<c:upBars>');
    expect(xml).toContain('<c:downBars>');
    expect(xml).toContain('<a:srgbClr val="22AA22"');
    expect(xml).toContain('<a:srgbClr val="AA2222"');
  });

  it('round-trips detailed hiLowLines + upDownBars', () => {
    const space = stockSpace(
      { spPr: { fill: makeSolidFill(makeColor(makeSrgbColor('888888'))) } },
      {
        gapWidth: 100,
        upBars: { spPr: { fill: makeSolidFill(makeColor(makeSrgbColor('22AA22'))) } },
        downBars: {},
      },
    );
    const back = parseChartXml(chartToBytes(space));
    const stock = back.plotArea.chart as StockChart;
    expect(typeof stock.hiLowLines).toBe('object');
    expect((stock.hiLowLines as { spPr?: unknown }).spPr).toBeDefined();
    expect(typeof stock.upDownBars).toBe('object');
    const ud = stock.upDownBars as { gapWidth?: number; upBars?: { spPr?: unknown }; downBars?: unknown };
    expect(ud.gapWidth).toBe(100);
    expect(ud.upBars?.spPr).toBeDefined();
    expect(ud.downBars).toBeDefined();
  });

  it('round-trips boolean true form when no detail is supplied', () => {
    const back = parseChartXml(chartToBytes(stockSpace(true, true)));
    const stock = back.plotArea.chart as StockChart;
    expect(stock.hiLowLines).toBe(true);
    expect(stock.upDownBars).toBe(true);
  });
});
