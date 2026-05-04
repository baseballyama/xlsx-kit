import { describe, expect, it } from 'vitest';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart';
import { makeWaterfallChart } from '../../src/chart/cx/chartex';
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

describe('chartex workbook integration', () => {
  it('emits CHARTEX_TYPE override and round-trips a Waterfall chart', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const cxSpace = makeWaterfallChart({
      catRef: 'Sheet1!$A$1:$A$5',
      valRef: 'Sheet1!$B$1:$B$5',
      subtotalIdx: [0, 4],
    });
    ws.drawing = makeDrawing([makeChartDrawingItem(makeTwoCellAnchor({ from: 'D2', to: 'J20' }), { cxSpace })]);

    const bytes = await workbookToBytes(wb);
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    expect(entries['xl/charts/chart1.xml']).toBeDefined();
    const ct = new TextDecoder().decode(entries['[Content_Types].xml']);
    expect(ct).toContain('vnd.ms-office.chartex+xml');
    expect(ct).not.toContain('drawingml.chart+xml');

    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const item = ws2.drawing?.items[0];
    expect(item?.content.kind).toBe('chart');
    if (item?.content.kind === 'chart') {
      const back = item.content.chart.cxSpace;
      if (!back) throw new Error('expected cxSpace');
      expect(back.kind).toBe('cxChartSpace');
      const series = back.chart.plotArea.series[0];
      expect(series?.layoutId).toBe('waterfall');
      expect(series?.layoutPr).toEqual({ kind: 'waterfall', subtotalIdx: [0, 4] });
      // The legacy `space` slot must be left undefined when cxSpace is set.
      expect(item.content.chart.space).toBeUndefined();
    }
  });

  it('mixed legacy + chartex on different sheets get distinct content types', async () => {
    const wb = createWorkbook();
    const sheetA = addWorksheet(wb, 'A');
    const sheetB = addWorksheet(wb, 'B');
    const legacy = makeChartSpace({
      plotArea: {
        chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A2' } })] }),
        catAx: { axId: 1, crossAx: 2 },
        valAx: { axId: 2, crossAx: 1 },
      },
    });
    const cxSpace = makeWaterfallChart({ catRef: 'A1:A3', valRef: 'B1:B3' });
    sheetA.drawing = makeDrawing([
      makeChartDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'B5' }), { space: legacy }),
    ]);
    sheetB.drawing = makeDrawing([makeChartDrawingItem(makeTwoCellAnchor({ from: 'A1', to: 'B5' }), { cxSpace })]);

    const bytes = await workbookToBytes(wb);
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    const ct = new TextDecoder().decode(entries['[Content_Types].xml']);
    expect(ct).toContain('drawingml.chart+xml');
    expect(ct).toContain('vnd.ms-office.chartex+xml');

    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const a = expectSheet(wb2.sheets[0]?.sheet);
    const b = expectSheet(wb2.sheets[1]?.sheet);
    const aItem = a.drawing?.items[0];
    const bItem = b.drawing?.items[0];
    if (aItem?.content.kind !== 'chart' || bItem?.content.kind !== 'chart') {
      throw new Error('expected chart items');
    }
    expect(aItem.content.chart.space).toBeDefined();
    expect(aItem.content.chart.cxSpace).toBeUndefined();
    expect(bItem.content.chart.cxSpace).toBeDefined();
    expect(bItem.content.chart.space).toBeUndefined();
  });
});
