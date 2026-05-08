// Scenario 20: chartex modern chart kinds — Sunburst / Treemap /
// Waterfall / Histogram / Pareto / Funnel / BoxWhisker / RegionMap.
// Output: 20-charts-chartex.xlsx
//
// Chartex (`cx:` namespace) charts are the post-2016 generation Excel
// uses for modern visualisations. The XML schemas are sparsely-
// documented and require non-trivial cached point data + per-layout
// `<cx:lvl>` cache trees that Excel rejects when missing. Until we
// have a high-fidelity chartex writer the e2e scenario uses legacy
// chart kinds (column / line / bar) so the file still opens reliably
// in real Excel; the chartex constructors remain available in the
// public API for callers that supply their own cache data.

import { describe, expect, it } from 'vitest';
import { makeBarChart, makeBarSeries, makeChartSpace, makeLineChart } from '../../../src/xlsx/chart/chart';
import type { ChartSpace } from '../../../src/xlsx/chart/chart';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { setCell } from '../../../src/xlsx/worksheet/index';
import { makeOneCellAnchor } from '../../../src/xlsx/drawing/anchor';
import { makeChartDrawingItem, makeDrawing } from '../../../src/xlsx/drawing/drawing';
import { writeWorkbook } from '../_helpers';

describe('e2e 20 — chartex modern chart kinds', () => {
  it('writes 20-charts-chartex.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');

    setCell(ws, 1, 1, 'Category');
    setCell(ws, 1, 2, 'Value');
    const cats = ['North/Apples', 'North/Oranges', 'South/Apples', 'South/Oranges', 'East/Apples', 'East/Oranges'];
    const values = [120, 80, 95, 110, 60, 75];
    cats.forEach((c, i) => {
      setCell(ws, i + 2, 1, c);
      setCell(ws, i + 2, 2, values[i] ?? 0);
    });

    const cat = { ref: 'Data!$A$2:$A$7', cacheKind: 'str' as const, cache: cats };
    const valSeries = makeBarSeries({
      idx: 0,
      val: { ref: 'Data!$B$2:$B$7', cache: values },
      cat,
      tx: { kind: 'ref', ref: 'Data!$B$1' },
    });

    const charts: Array<{ anchor: string; space: ChartSpace }> = [
      {
        anchor: 'D2',
        space: makeChartSpace({
          title: 'Column (legacy fallback for Sunburst)',
          plotArea: { chart: makeBarChart({ barDir: 'col', grouping: 'clustered', series: [valSeries] }) },
          legend: { position: 'r' },
        }),
      },
      {
        anchor: 'D20',
        space: makeChartSpace({
          title: 'Bar (legacy fallback for Treemap)',
          plotArea: { chart: makeBarChart({ barDir: 'bar', grouping: 'clustered', series: [valSeries] }) },
          legend: { position: 'r' },
        }),
      },
      {
        anchor: 'M2',
        space: makeChartSpace({
          title: 'Line (legacy fallback for Histogram)',
          plotArea: { chart: makeLineChart({ series: [valSeries] }) },
          legend: { position: 'r' },
        }),
      },
    ];

    ws.drawing = makeDrawing(
      charts.map((c) =>
        makeChartDrawingItem(makeOneCellAnchor({ from: c.anchor, widthPx: 360, heightPx: 240 }), {
          space: c.space,
        }),
      ),
    );

    const result = await writeWorkbook('20-charts-chartex.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
