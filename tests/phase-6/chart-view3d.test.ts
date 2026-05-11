import { describe, expect, it } from 'vitest';
import { makeBar3DChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';
import { makeColor, makeSrgbColor } from '../../src/drawing/dml/colors';
import { makeSolidFill } from '../../src/drawing/dml/fill';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

const bar3DSpace = (extra: Partial<Parameters<typeof makeChartSpace>[0]> = {}) =>
  makeChartSpace({
    plotArea: {
      chart: makeBar3DChart({
        series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } })],
        axIds: [1, 2, 3],
      }),
      catAx: { axId: 1, crossAx: 2 },
      valAx: { axId: 2, crossAx: 1 },
    },
    ...extra,
  });

describe('view3D + floor/sideWall/backWall', () => {
  it('serialises view3D between autoTitleDeleted and plotArea', () => {
    const xml = decode(
      chartToBytes(
        bar3DSpace({ view3D: { rotX: 15, rotY: 20, depthPercent: 100, rAngAx: true, perspective: 30 } }),
      ),
    );
    expect(xml).toContain(
      '<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:depthPercent val="100"/><c:rAngAx val="1"/><c:perspective val="30"/></c:view3D>',
    );
    const view3DPos = xml.indexOf('<c:view3D>');
    const plotAreaPos = xml.indexOf('<c:plotArea>');
    expect(view3DPos).toBeGreaterThan(0);
    expect(view3DPos).toBeLessThan(plotAreaPos);
  });

  it('round-trips view3D through parseChartXml', () => {
    const back = parseChartXml(
      chartToBytes(bar3DSpace({ view3D: { rotX: -10, rotY: 45, perspective: 0 } })),
    );
    expect(back.view3D).toEqual({ rotX: -10, rotY: 45, perspective: 0 });
  });

  it('serialises and round-trips floor / sideWall / backWall with spPr', () => {
    const fill = makeSolidFill(makeColor(makeSrgbColor('EEEEEE')));
    const back = parseChartXml(
      chartToBytes(
        bar3DSpace({
          floor: { thickness: 0, spPr: { fill } },
          sideWall: { thickness: 0, spPr: { fill } },
          backWall: { thickness: 0, spPr: { fill } },
        }),
      ),
    );
    expect(back.floor?.thickness).toBe(0);
    expect(back.sideWall?.thickness).toBe(0);
    expect(back.backWall?.thickness).toBe(0);
    expect(back.floor?.spPr?.fill).toBeDefined();
  });

  it('omits view3D / walls when unset', () => {
    const xml = decode(chartToBytes(bar3DSpace()));
    expect(xml).not.toContain('<c:view3D>');
    expect(xml).not.toContain('<c:floor>');
    expect(xml).not.toContain('<c:sideWall>');
    expect(xml).not.toContain('<c:backWall>');
  });
});
