import { describe, expect, it } from 'vitest';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

describe('issue #48 — ChartSpace.style is exposed and round-trips', () => {
  const minimalSpace = (style?: number) =>
    makeChartSpace({
      plotArea: {
        chart: makeBarChart({
          series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } })],
        }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
      ...(style !== undefined ? { style } : {}),
    });

  it('serialises <c:style> between roundedCorners and <c:chart>', () => {
    const xml = decode(chartToBytes(minimalSpace(10)));
    expect(xml).toContain('<c:roundedCorners val="0"/><c:style val="10"/><c:chart>');
  });

  it('omits <c:style> when style is unset', () => {
    const xml = decode(chartToBytes(minimalSpace()));
    expect(xml).not.toContain('<c:style');
  });

  it('round-trips style through parseChartXml', () => {
    const back = parseChartXml(chartToBytes(minimalSpace(42)));
    expect(back.style).toBe(42);
  });
});
