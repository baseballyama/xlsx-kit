import { describe, expect, it } from 'vitest';
import {
  type BarChart,
  makeBarChart,
  makeBarSeries,
  makeChartSpace,
} from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

const space = (
  catAx?: Parameters<typeof makeChartSpace>[0]['plotArea']['catAx'],
  valAx?: Parameters<typeof makeChartSpace>[0]['plotArea']['valAx'],
) =>
  makeChartSpace({
    plotArea: {
      chart: makeBarChart({
        series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } })],
      }) as BarChart,
      catAx: catAx ?? { axId: 1, crossAx: 2 },
      valAx: valAx ?? { axId: 2, crossAx: 1 },
    },
  });

describe('issue #46 — axis scaling/crosses + extended axis attrs round-trip', () => {
  it('serialises scaling.orientation = maxMin (horizontal-bar reversal)', () => {
    const xml = decode(
      chartToBytes(space({ axId: 1, crossAx: 2, scaling: { orientation: 'maxMin' } })),
    );
    expect(xml).toContain('<c:scaling><c:orientation val="maxMin"/></c:scaling>');
  });

  it('serialises scaling.max for 100%-stacked value axis cap', () => {
    const xml = decode(chartToBytes(space(undefined, { axId: 2, crossAx: 1, scaling: { max: 1 } })));
    expect(xml).toContain('<c:max val="1"/>');
  });

  it('serialises crosses = max (cross at right edge)', () => {
    const xml = decode(chartToBytes(space({ axId: 1, crossAx: 2, crosses: 'max' })));
    expect(xml).toContain('<c:crosses val="max"/>');
  });

  it('serialises crossesAt instead of crosses when both present', () => {
    const xml = decode(
      chartToBytes(space({ axId: 1, crossAx: 2, crosses: 'autoZero', crossesAt: 5 })),
    );
    expect(xml).toContain('<c:crossesAt val="5"/>');
    // The catAx section uses crossesAt; only the valAx default <c:crosses> remains.
    const catSection = xml.slice(xml.indexOf('<c:catAx>'), xml.indexOf('</c:catAx>'));
    expect(catSection).not.toContain('<c:crosses ');
  });

  it('serialises axis numFmt override', () => {
    const xml = decode(
      chartToBytes(
        space(undefined, {
          axId: 2,
          crossAx: 1,
          numFmt: { formatCode: '#,##0', sourceLinked: false },
        }),
      ),
    );
    expect(xml).toContain('<c:numFmt formatCode="#,##0" sourceLinked="0"/>');
  });

  it('serialises majorTickMark / minorTickMark / tickLblPos overrides', () => {
    const xml = decode(
      chartToBytes(
        space({
          axId: 1,
          crossAx: 2,
          majorTickMark: 'cross',
          minorTickMark: 'in',
          tickLblPos: 'low',
        }),
      ),
    );
    expect(xml).toContain('<c:majorTickMark val="cross"/>');
    expect(xml).toContain('<c:minorTickMark val="in"/>');
    expect(xml).toContain('<c:tickLblPos val="low"/>');
  });

  it('serialises valAx majorUnit / minorUnit / crossBetween', () => {
    const xml = decode(
      chartToBytes(
        space(undefined, {
          axId: 2,
          crossAx: 1,
          majorUnit: 50,
          minorUnit: 10,
          crossBetween: 'midCat',
        }),
      ),
    );
    expect(xml).toContain('<c:crossBetween val="midCat"/>');
    expect(xml).toContain('<c:majorUnit val="50"/>');
    expect(xml).toContain('<c:minorUnit val="10"/>');
  });

  it('serialises catAx auto = 0, lblAlgn, lblOffset, noMultiLvlLbl overrides', () => {
    const xml = decode(
      chartToBytes(
        space({
          axId: 1,
          crossAx: 2,
          auto: false,
          lblAlgn: 'l',
          lblOffset: 50,
          noMultiLvlLbl: true,
        }),
      ),
    );
    expect(xml).toContain('<c:auto val="0"/>');
    expect(xml).toContain('<c:lblAlgn val="l"/>');
    expect(xml).toContain('<c:lblOffset val="50"/>');
    expect(xml).toContain('<c:noMultiLvlLbl val="1"/>');
  });

  it('round-trips a full set of axis options', () => {
    const original = space(
      {
        axId: 1,
        crossAx: 2,
        scaling: { orientation: 'maxMin' },
        crosses: 'max',
        majorTickMark: 'cross',
        tickLblPos: 'low',
        auto: false,
        lblAlgn: 'r',
        lblOffset: 80,
      },
      {
        axId: 2,
        crossAx: 1,
        scaling: { max: 1, min: 0 },
        crossBetween: 'midCat',
        majorUnit: 0.25,
        numFmt: { formatCode: '0.00%', sourceLinked: false },
      },
    );
    const back = parseChartXml(chartToBytes(original));
    expect(back.plotArea.catAx?.scaling).toEqual({ orientation: 'maxMin' });
    expect(back.plotArea.catAx?.crosses).toBe('max');
    expect(back.plotArea.catAx?.majorTickMark).toBe('cross');
    expect(back.plotArea.catAx?.tickLblPos).toBe('low');
    expect(back.plotArea.catAx?.auto).toBe(false);
    expect(back.plotArea.catAx?.lblAlgn).toBe('r');
    expect(back.plotArea.catAx?.lblOffset).toBe(80);
    expect(back.plotArea.valAx?.scaling).toEqual({ max: 1, min: 0 });
    expect(back.plotArea.valAx?.crossBetween).toBe('midCat');
    expect(back.plotArea.valAx?.majorUnit).toBe(0.25);
    expect(back.plotArea.valAx?.numFmt).toEqual({ formatCode: '0.00%', sourceLinked: false });
  });

  it('preserves the default minMax orientation when scaling is unset', () => {
    const xml = decode(chartToBytes(space()));
    expect(xml).toContain('<c:scaling><c:orientation val="minMax"/></c:scaling>');
  });
});
