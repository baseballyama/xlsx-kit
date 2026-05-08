// Scenario 13: a bar chart on a worksheet — references real cell
// values via NumericRef strings.
// Output: 13-chart-bar.xlsx
//
// What to verify in Excel:
// - Sheet "Data" has labels in A2:A4 (Q1/Q2/Q3) and values in
//   B2:B4 (10/30/20).
// - Sheet "Data" also has a clustered column chart titled "Quarterly
//   Sales" anchored from B5:H20 (or thereabouts) showing the three
//   bars with category labels Q1/Q2/Q3.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../../src/workbook/index';
import { setCell } from '../../../src/worksheet/index';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../../src/chart/chart';
import { makeAbsoluteAnchor, makeOneCellAnchor } from '../../../src/drawing/anchor';
import { makeChartDrawingItem, makeDrawing } from '../../../src/drawing/drawing';
import { writeWorkbook } from '../_helpers';

describe('e2e 13 — bar chart', () => {
  it('writes 13-chart-bar.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');

    setCell(ws, 1, 1, 'Quarter');
    setCell(ws, 1, 2, 'Sales');
    const labels = ['Q1', 'Q2', 'Q3'];
    const values = [10, 30, 20];
    for (let i = 0; i < 3; i++) {
      setCell(ws, i + 2, 1, labels[i] ?? '');
      setCell(ws, i + 2, 2, values[i] ?? 0);
    }

    const series = makeBarSeries({
      idx: 0,
      val: { ref: 'Data!$B$2:$B$4', cache: values, formatCode: 'General' },
      cat: { ref: 'Data!$A$2:$A$4', cacheKind: 'str', cache: labels },
      tx: { kind: 'ref', ref: 'Data!$B$1' },
    });
    const space = makeChartSpace({
      title: 'Quarterly Sales',
      plotArea: {
        chart: makeBarChart({ barDir: 'col', grouping: 'clustered', series: [series] }),
      },
      legend: { position: 'r' },
      plotVisOnly: true,
    });

    ws.drawing = makeDrawing([
      makeChartDrawingItem(
        makeOneCellAnchor({ from: 'D2', widthPx: 480, heightPx: 320 }),
        { space },
      ),
    ]);
    void makeAbsoluteAnchor;

    const result = await writeWorkbook('13-chart-bar.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
