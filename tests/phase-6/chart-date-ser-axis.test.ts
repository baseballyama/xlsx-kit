import { describe, expect, it } from 'vitest';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

describe('DateAxis + SeriesAxis', () => {
  it('serialises and round-trips a date axis', () => {
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({
          series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A3' } })],
        }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
        dateAx: {
          axId: 3,
          crossAx: 2,
          baseTimeUnit: 'months',
          majorUnit: 1,
          majorTimeUnit: 'years',
          minorUnit: 1,
          minorTimeUnit: 'months',
        },
      },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).toContain('<c:dateAx>');
    expect(xml).toContain('<c:baseTimeUnit val="months"/>');
    expect(xml).toContain('<c:majorTimeUnit val="years"/>');
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.dateAx?.axId).toBe(3);
    expect(back.plotArea.dateAx?.baseTimeUnit).toBe('months');
    expect(back.plotArea.dateAx?.majorTimeUnit).toBe('years');
    expect(back.plotArea.dateAx?.majorUnit).toBe(1);
    expect(back.plotArea.dateAx?.minorTimeUnit).toBe('months');
  });

  it('serialises and round-trips a series axis with tickLblSkip / tickMarkSkip', () => {
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({
          series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A3' } })],
        }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
        serAx: { axId: 4, crossAx: 1, tickLblSkip: 2, tickMarkSkip: 1 },
      },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).toContain('<c:serAx>');
    expect(xml).toContain('<c:tickLblSkip val="2"/>');
    expect(xml).toContain('<c:tickMarkSkip val="1"/>');
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.serAx?.tickLblSkip).toBe(2);
    expect(back.plotArea.serAx?.tickMarkSkip).toBe(1);
  });

  it('omits both axes when unset', () => {
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1' } })] }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).not.toContain('<c:dateAx>');
    expect(xml).not.toContain('<c:serAx>');
  });
});
