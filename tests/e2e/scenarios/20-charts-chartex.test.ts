// Scenario 20: chartex modern chart kinds — Sunburst / Treemap /
// Waterfall / Histogram / Pareto / Funnel / BoxWhisker / RegionMap.
// Output: 20-charts-chartex.xlsx
//
// Chartex (`cx:` namespace) charts are the post-2016 generation Excel
// uses for modern visualisations. Each has its own data model — most
// take a category column + value column.
//
// What to verify in Excel:
// - "Data" sheet has hierarchical-ish categories (regions / products)
//   with a numeric value column.
// - 8 charts are anchored across the sheet, each rendering its
//   chartex shape (Sunburst, Treemap, etc.). Excel 2016+ required;
//   older Excel will render a placeholder or refuse the cx: namespace.

import { describe, expect, it } from 'vitest';
import {
  makeBoxWhiskerChart,
  makeFunnelChart,
  makeHistogramChart,
  makeParetoChart,
  makeRegionMapChart,
  makeSunburstChart,
  makeTreemapChart,
  makeWaterfallChart,
} from '../../../src/chart/cx/chartex';
import type { CxChartSpace } from '../../../src/chart/cx/chartex';
import { addWorksheet, createWorkbook, setCell } from '../../../src/index';
import { makeOneCellAnchor } from '../../../src/drawing/anchor';
import { makeChartDrawingItem, makeDrawing } from '../../../src/drawing/drawing';
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

    const charts: Array<{ anchor: string; space: CxChartSpace; title: string }> = [
      { anchor: 'D2', title: 'Sunburst', space: makeSunburstChart({ catRef: 'Data!$A$2:$A$7', valRef: 'Data!$B$2:$B$7' }) },
      { anchor: 'D20', title: 'Treemap', space: makeTreemapChart({ catRef: 'Data!$A$2:$A$7', valRef: 'Data!$B$2:$B$7' }) },
      { anchor: 'D38', title: 'Waterfall', space: makeWaterfallChart({ catRef: 'Data!$A$2:$A$7', valRef: 'Data!$B$2:$B$7', subtotalIdx: [3] }) },
      { anchor: 'M2', title: 'Histogram', space: makeHistogramChart({ valRef: 'Data!$B$2:$B$7' }) },
      { anchor: 'M20', title: 'Pareto', space: makeParetoChart({ catRef: 'Data!$A$2:$A$7', valRef: 'Data!$B$2:$B$7' }) },
      { anchor: 'M38', title: 'Funnel', space: makeFunnelChart({ catRef: 'Data!$A$2:$A$7', valRef: 'Data!$B$2:$B$7' }) },
      { anchor: 'V2', title: 'BoxWhisker', space: makeBoxWhiskerChart({ valRef: 'Data!$B$2:$B$7' }) },
      { anchor: 'V20', title: 'RegionMap', space: makeRegionMapChart({ catRef: 'Data!$A$2:$A$7', valRef: 'Data!$B$2:$B$7' }) },
    ];

    void charts.reduce((_, c) => c.title.length, 0); // hush unused-cap

    ws.drawing = makeDrawing(
      charts.map((c) =>
        makeChartDrawingItem(makeOneCellAnchor({ from: c.anchor, widthPx: 360, heightPx: 240 }), {
          cxSpace: c.space,
        }),
      ),
    );

    const result = await writeWorkbook('20-charts-chartex.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
