import { describe, expect, it } from 'vitest';
import type { Gridlines, ValueAxis } from '../../src/chart';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';
import { makeColor, makeSrgbColor } from '../../src/drawing/dml/colors';
import { makeSolidFill } from '../../src/drawing/dml/fill';
import { makeShapeProperties } from '../../src/drawing/dml/shape-properties';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

describe('issue #57 — majorGridlines / minorGridlines accept the rich `{ spPr }` form', () => {
  it('keeps the boolean `true` form emitting a self-closing element (back-compat)', () => {
    const xml = decode(
      chartToBytes(
        makeChartSpace({
          plotArea: {
            chart: makeBarChart({
              series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } })],
            }),
            catAx: { axId: 1, crossAx: 2 },
            valAx: { axId: 2, crossAx: 1, majorGridlines: true, minorGridlines: true },
          },
        }),
      ),
    );
    expect(xml).toContain('<c:majorGridlines/>');
    expect(xml).toContain('<c:minorGridlines/>');
  });

  it('serialises a styled gridline with <c:spPr> when given the object form', () => {
    const grid: Gridlines = {
      spPr: makeShapeProperties({
        ln: { fill: makeSolidFill(makeColor(makeSrgbColor('D9D9D9'))) },
      }),
    };
    const valAx: ValueAxis = {
      axId: 2,
      crossAx: 1,
      majorGridlines: grid,
    };
    const xml = decode(
      chartToBytes(
        makeChartSpace({
          plotArea: {
            chart: makeBarChart({
              series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } })],
            }),
            catAx: { axId: 1, crossAx: 2 },
            valAx,
          },
        }),
      ),
    );
    expect(xml).toContain('<c:majorGridlines><c:spPr');
    expect(xml).toContain('<a:srgbClr val="D9D9D9">');
    expect(xml).toContain('</c:majorGridlines>');
  });

  it('round-trips both forms through parseChartXml', () => {
    const grid: Gridlines = {
      spPr: makeShapeProperties({
        ln: { fill: makeSolidFill(makeColor(makeSrgbColor('AAAAAA'))) },
      }),
    };
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({
          series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } })],
        }),
        catAx: { axId: 1, crossAx: 2, minorGridlines: true },
        valAx: { axId: 2, crossAx: 1, majorGridlines: grid },
      },
    });
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.catAx?.minorGridlines).toBe(true);
    const major = back.plotArea.valAx?.majorGridlines;
    expect(major).not.toBe(true);
    expect(typeof major === 'object' && major?.spPr?.ln?.fill?.kind).toBe('solidFill');
  });
});
