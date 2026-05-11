import { describe, expect, it } from 'vitest';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

describe('manualLayout for title / plotArea / legend', () => {
  it('serialises and round-trips title.layout', () => {
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A3' } })] }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
      title: {
        text: 'Custom Title',
        layout: {
          manualLayout: {
            layoutTarget: 'inner',
            xMode: 'edge',
            yMode: 'edge',
            x: 0.1,
            y: 0.05,
            w: 0.6,
            h: 0.1,
          },
        },
      },
    });
    const xml = decode(chartToBytes(space));
    expect(xml).toContain('<c:manualLayout>');
    expect(xml).toContain('<c:layoutTarget val="inner"/>');
    expect(xml).toContain('<c:x val="0.1"/>');
    const back = parseChartXml(chartToBytes(space));
    expect(back.title?.layout?.manualLayout?.layoutTarget).toBe('inner');
    expect(back.title?.layout?.manualLayout?.x).toBe(0.1);
    expect(back.title?.layout?.manualLayout?.w).toBe(0.6);
  });

  it('serialises and round-trips plotArea.layout', () => {
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A3' } })] }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
        layout: {
          manualLayout: { x: 0.1, y: 0.15, w: 0.8, h: 0.7 },
        },
      },
    });
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.layout?.manualLayout?.x).toBe(0.1);
    expect(back.plotArea.layout?.manualLayout?.h).toBe(0.7);
  });

  it('serialises and round-trips legend.layout', () => {
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A3' } })] }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
      legend: {
        position: 'r',
        layout: { manualLayout: { x: 0.8, y: 0.1, w: 0.18, h: 0.6 } },
      },
    });
    const back = parseChartXml(chartToBytes(space));
    expect(back.legend?.layout?.manualLayout?.x).toBe(0.8);
    expect(back.legend?.layout?.manualLayout?.w).toBe(0.18);
  });

  it('keeps emitting empty <c:layout/> when layout is unset (output unchanged)', () => {
    const xml = decode(
      chartToBytes(
        makeChartSpace({
          plotArea: {
            chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1' } })] }),
            catAx: { axId: 1, crossAx: 2 },
            valAx: { axId: 2, crossAx: 1 },
          },
        }),
      ),
    );
    expect(xml).toContain('<c:layout/>');
    expect(xml).not.toContain('<c:manualLayout>');
  });
});
