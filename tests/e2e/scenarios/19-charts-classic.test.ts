// Scenario 19: classic ChartML chart kinds — line / area / pie / doughnut
// / scatter / radar — all anchored on a single sheet of source data.
// Output: 19-charts-classic.xlsx
//
// What to verify in Excel:
// - "Data" sheet rows 1..6 carry months Jan..May with three series A/B/C.
// - Six charts are anchored across the sheet (D2, D20, D38, K2, K20, K38)
//   showing the same data via different chart types. Each should render
//   correctly with axis labels / legend visible.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook, setCell } from '../../../src/index';
import {
  makeAreaChart,
  makeBarSeries,
  makeChartSpace,
  makeDoughnutChart,
  makeLineChart,
  makePieChart,
  makeRadarChart,
  makeScatterChart,
  makeScatterSeries,
} from '../../../src/chart/chart';
import { makeOneCellAnchor } from '../../../src/drawing/anchor';
import { makeChartDrawingItem, makeDrawing } from '../../../src/drawing/drawing';
import type { ChartReference } from '../../../src/drawing/drawing';
import type { ChartSpace } from '../../../src/chart/chart';
import { writeWorkbook } from '../_helpers';

describe('e2e 19 — classic chart kinds', () => {
  it('writes 19-charts-classic.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');

    setCell(ws, 1, 1, 'Month');
    setCell(ws, 1, 2, 'A');
    setCell(ws, 1, 3, 'B');
    setCell(ws, 1, 4, 'C');
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May'];
    const seriesA = [10, 30, 20, 40, 25];
    const seriesB = [20, 25, 35, 30, 45];
    const seriesC = [15, 18, 22, 28, 32];
    months.forEach((m, i) => {
      setCell(ws, i + 2, 1, m);
      setCell(ws, i + 2, 2, seriesA[i] ?? 0);
      setCell(ws, i + 2, 3, seriesB[i] ?? 0);
      setCell(ws, i + 2, 4, seriesC[i] ?? 0);
    });

    const cat = { ref: 'Data!$A$2:$A$6', cacheKind: 'str' as const, cache: months };
    const seriesABC = [
      makeBarSeries({ idx: 0, val: { ref: 'Data!$B$2:$B$6', cache: seriesA }, cat, tx: { kind: 'ref', ref: 'Data!$B$1' } }),
      makeBarSeries({ idx: 1, val: { ref: 'Data!$C$2:$C$6', cache: seriesB }, cat, tx: { kind: 'ref', ref: 'Data!$C$1' } }),
      makeBarSeries({ idx: 2, val: { ref: 'Data!$D$2:$D$6', cache: seriesC }, cat, tx: { kind: 'ref', ref: 'Data!$D$1' } }),
    ];

    const charts: Array<{ anchor: string; space: ChartSpace }> = [
      {
        anchor: 'F2',
        space: makeChartSpace({ title: 'Line', plotArea: { chart: makeLineChart({ series: seriesABC }) }, legend: { position: 'r' } }),
      },
      {
        anchor: 'F20',
        space: makeChartSpace({ title: 'Area (stacked)', plotArea: { chart: makeAreaChart({ grouping: 'stacked', series: seriesABC }) }, legend: { position: 'r' } }),
      },
      {
        anchor: 'F38',
        space: makeChartSpace({ title: 'Pie (series A only)', plotArea: { chart: makePieChart({ series: [seriesABC[0]!] }) }, legend: { position: 'b' } }),
      },
      {
        anchor: 'O2',
        space: makeChartSpace({ title: 'Doughnut', plotArea: { chart: makeDoughnutChart({ series: [seriesABC[0]!], holeSize: 50 }) }, legend: { position: 'b' } }),
      },
      {
        anchor: 'O20',
        space: makeChartSpace({
          title: 'Scatter',
          plotArea: {
            chart: makeScatterChart({
              scatterStyle: 'lineMarker',
              series: [
                makeScatterSeries({
                  idx: 0,
                  xVal: { ref: 'Data!$B$2:$B$6', cache: seriesA },
                  yVal: { ref: 'Data!$C$2:$C$6', cache: seriesB },
                  tx: { kind: 'literal', value: 'A vs B' },
                }),
              ],
            }),
          },
          legend: { position: 'r' },
        }),
      },
      {
        anchor: 'O38',
        space: makeChartSpace({
          title: 'Radar',
          plotArea: { chart: makeRadarChart({ radarStyle: 'standard', series: seriesABC }) },
          legend: { position: 'r' },
        }),
      },
    ];

    const items = charts.map((c) =>
      makeChartDrawingItem(makeOneCellAnchor({ from: c.anchor, widthPx: 360, heightPx: 240 }), {
        space: c.space,
      } satisfies ChartReference),
    );
    ws.drawing = makeDrawing(items);

    const result = await writeWorkbook('19-charts-classic.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
