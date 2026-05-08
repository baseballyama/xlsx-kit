// Scenario 21: chart series decorations — data labels, trendline,
// error bars. All on a single bar/scatter chart so the user can see
// each at once.
// Output: 21-chart-decorations.xlsx
//
// What to verify in Excel:
// - Top chart (column): each bar shows its value as a data label
//   above the bar. A linear trendline cuts through the series.
// - Bottom chart (scatter): error bars on the Y axis show ±10% per
//   point, with a separate exponential trendline.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { setCell } from '../../../src/xlsx/worksheet/index';
import {
  makeBarChart,
  makeBarSeries,
  makeChartSpace,
  makeScatterChart,
  makeScatterSeries,
} from '../../../src/xlsx/chart/chart';
import type { ChartSpace } from '../../../src/xlsx/chart/chart';
import { makeOneCellAnchor } from '../../../src/xlsx/drawing/anchor';
import { makeChartDrawingItem, makeDrawing } from '../../../src/xlsx/drawing/drawing';
import { writeWorkbook } from '../_helpers';

describe('e2e 21 — chart decorations (dLbls / trendline / errBars)', () => {
  it('writes 21-chart-decorations.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');

    setCell(ws, 1, 1, 'Period');
    setCell(ws, 1, 2, 'Sales');
    setCell(ws, 1, 3, 'X');
    setCell(ws, 1, 4, 'Y');
    const periods = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul'];
    const sales = [100, 130, 150, 170, 200, 220, 250];
    const xs = [1, 2, 3, 4, 5, 6, 7];
    const ys = [2.7, 7.4, 20.1, 54.6, 148, 403, 1097];
    periods.forEach((p, i) => {
      setCell(ws, i + 2, 1, p);
      setCell(ws, i + 2, 2, sales[i] ?? 0);
      setCell(ws, i + 2, 3, xs[i] ?? 0);
      setCell(ws, i + 2, 4, ys[i] ?? 0);
    });

    // Bar chart with data labels + linear trendline
    const barSeries = makeBarSeries({
      idx: 0,
      val: { ref: 'Data!$B$2:$B$8', cache: sales },
      cat: { ref: 'Data!$A$2:$A$8', cacheKind: 'str', cache: periods },
      tx: { kind: 'ref', ref: 'Data!$B$1' },
    });
    barSeries.dLbls = { showVal: true };
    barSeries.trendline = [{ trendlineType: 'linear', dispEq: true, dispRSqr: true }];
    const barSpace: ChartSpace = makeChartSpace({
      title: 'Sales with linear trendline',
      plotArea: { chart: makeBarChart({ barDir: 'col', grouping: 'clustered', series: [barSeries] }) },
      legend: { position: 'r' },
    });

    // Scatter with errBars (Y-axis ±10%) + exponential trendline
    const scatterSeries = makeScatterSeries({
      idx: 0,
      xVal: { ref: 'Data!$C$2:$C$8', cache: xs },
      yVal: { ref: 'Data!$D$2:$D$8', cache: ys },
      tx: { kind: 'literal', value: 'Exponential growth' },
    });
    scatterSeries.errBars = [
      {
        errDir: 'y',
        errBarType: 'both',
        errValType: 'percentage',
        val: 10,
        noEndCap: false,
      },
    ];
    scatterSeries.trendline = [{ trendlineType: 'exp', dispEq: true }];
    const scatterSpace: ChartSpace = makeChartSpace({
      title: 'Y vs X with ±10% error + exp trendline',
      plotArea: { chart: makeScatterChart({ scatterStyle: 'lineMarker', series: [scatterSeries] }) },
      legend: { position: 'r' },
    });

    ws.drawing = makeDrawing([
      makeChartDrawingItem(makeOneCellAnchor({ from: 'F2', widthPx: 480, heightPx: 320 }), { space: barSpace }),
      makeChartDrawingItem(makeOneCellAnchor({ from: 'F22', widthPx: 480, heightPx: 320 }), { space: scatterSpace }),
    ]);

    const result = await writeWorkbook('21-chart-decorations.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
