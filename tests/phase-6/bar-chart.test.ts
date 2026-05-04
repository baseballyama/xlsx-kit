import { describe, expect, it } from 'vitest';
import { type ChartSpace, makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml';
import { makeTwoCellAnchor } from '../../src/drawing/anchor';
import { makeChartDrawingItem, makeDrawing } from '../../src/drawing/drawing';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import type { Worksheet } from '../../src/worksheet/worksheet';

const expectSheet = (ws: Worksheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  return ws;
};

describe('chart-xml round-trip — BarChart', () => {
  it('round-trips a clustered column chart with one series + cached cat/val', () => {
    const chart = makeBarChart({
      barDir: 'col',
      grouping: 'clustered',
      series: [
        makeBarSeries({
          idx: 0,
          val: { ref: 'Sheet1!$B$2:$B$5', cache: [10, 20, 30, 40] },
          cat: {
            ref: 'Sheet1!$A$2:$A$5',
            cacheKind: 'str',
            cache: ['Q1', 'Q2', 'Q3', 'Q4'],
          },
          tx: { kind: 'literal', value: 'Sales' },
        }),
      ],
    });
    const space = makeChartSpace({
      title: 'Quarterly Sales',
      legend: { position: 'r' },
      plotArea: {
        chart,
        catAx: { axId: 1, crossAx: 2, position: 'b' },
        valAx: { axId: 2, crossAx: 1, position: 'l', majorGridlines: true },
      },
    });
    const back = parseChartXml(chartToBytes(space));
    expect(back.title).toBe('Quarterly Sales');
    expect(back.legend?.position).toBe('r');
    const bar = back.plotArea.chart;
    if (bar.kind !== 'bar') throw new Error('expected bar chart');
    expect(bar.barDir).toBe('col');
    expect(bar.grouping).toBe('clustered');
    expect(bar.series.length).toBe(1);
    const s = bar.series[0];
    expect(s?.val.ref).toBe('Sheet1!$B$2:$B$5');
    expect(s?.val.cache).toEqual([10, 20, 30, 40]);
    expect(s?.cat?.ref).toBe('Sheet1!$A$2:$A$5');
    expect(s?.cat?.cache).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
    expect(back.plotArea.catAx?.axId).toBe(1);
    expect(back.plotArea.valAx?.majorGridlines).toBe(true);
  });

  it('round-trips a horizontal bar chart with multiple series', () => {
    const chart = makeBarChart({
      barDir: 'bar',
      grouping: 'stacked',
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A3' } }), makeBarSeries({ idx: 1, val: { ref: 'B1:B3' } })],
    });
    const back = parseChartXml(
      chartToBytes(
        makeChartSpace({
          plotArea: {
            chart,
            catAx: { axId: 1, crossAx: 2 },
            valAx: { axId: 2, crossAx: 1 },
          },
        }),
      ),
    );
    const bar = back.plotArea.chart;
    if (bar.kind !== 'bar') throw new Error('expected bar chart');
    expect(bar.barDir).toBe('bar');
    expect(bar.grouping).toBe('stacked');
    expect(bar.series.length).toBe(2);
  });

  it('rejects a non-chartSpace root', () => {
    expect(() => parseChartXml('<foo/>')).toThrowError(/expected chartSpace/);
  });
});

describe('full chart round-trip through saveWorkbook → loadWorkbook', () => {
  it('emits chart part + drawing rels + manifest Override and reads them back', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Charts');
    const chart = makeBarChart({
      series: [
        makeBarSeries({
          idx: 0,
          val: { ref: 'Charts!$B$1:$B$3', cache: [1, 2, 3] },
        }),
      ],
    });
    const space = makeChartSpace({
      title: 'Test',
      plotArea: {
        chart,
        catAx: { axId: 1, crossAx: 2, position: 'b' },
        valAx: { axId: 2, crossAx: 1, position: 'l' },
      },
    });
    ws.drawing = makeDrawing([makeChartDrawingItem(makeTwoCellAnchor({ from: 'D2', to: 'J20' }), { space })]);

    const bytes = await workbookToBytes(wb);
    // Quick check: the unzipped manifest declares the chart override.
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    expect(entries['xl/charts/chart1.xml']).toBeDefined();
    expect(entries['xl/drawings/_rels/drawing1.xml.rels']).toBeDefined();
    const ct = new TextDecoder().decode(entries['[Content_Types].xml']);
    expect(ct).toContain('drawingml.chart+xml');

    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.drawing?.items.length).toBe(1);
    const item = ws2.drawing?.items[0];
    expect(item?.content.kind).toBe('chart');
    if (item?.content.kind === 'chart') {
      const back = item.content.chart.space as ChartSpace;
      expect(back.title).toBe('Test');
      const barBack = back.plotArea.chart;
      if (barBack.kind !== 'bar') throw new Error('expected bar chart');
      expect(barBack.barDir).toBe('col');
      expect(barBack.series[0]?.val.ref).toBe('Charts!$B$1:$B$3');
      expect(barBack.series[0]?.val.cache).toEqual([1, 2, 3]);
    }
  });

  it('handles multiple charts on multiple sheets with workbook-global chartN ids', async () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    const mk = (lbl: string) =>
      makeChartSpace({
        title: lbl,
        plotArea: {
          chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A2' } })] }),
          catAx: { axId: 1, crossAx: 2 },
          valAx: { axId: 2, crossAx: 1 },
        },
      });
    a.drawing = makeDrawing([
      makeChartDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'B5' }), { space: mk('chart-A1') }),
    ]);
    b.drawing = makeDrawing([
      makeChartDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'B5' }), { space: mk('chart-B1') }),
      makeChartDrawingItem(makeTwoCellAnchor({ from: 'C1', to: 'D5' }), { space: mk('chart-B2') }),
    ]);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const titles = (sheet: Worksheet | undefined): string[] =>
      (sheet?.drawing?.items ?? [])
        .map((i) => (i.content.kind === 'chart' ? i.content.chart.space?.title : undefined))
        .filter((t): t is string => t !== undefined);
    expect(titles(wb2.sheets[0]?.sheet)).toEqual(['chart-A1']);
    expect(titles(wb2.sheets[1]?.sheet)).toEqual(['chart-B1', 'chart-B2']);
  });
});
