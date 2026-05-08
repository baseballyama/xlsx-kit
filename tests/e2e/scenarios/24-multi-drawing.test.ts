// Scenario 24: multiple drawings on a single worksheet — bar chart +
// line chart + an embedded image, all in the same `xl/drawings/...`
// part. Output: 24-multi-drawing.xlsx
//
// What to verify in Excel:
// - One sheet "Combo" with three drawings:
//   * E2 → clustered bar chart "Quarterly Sales"
//   * E20 → line chart "Trend"
//   * N2 → tiny PNG image (the same fixture as scenario 18)
// - All three render together without "we found a problem" recovery.
// - The image stays positioned at N2 even after Excel re-saves the file.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../../src/workbook/index';
import { setCell } from '../../../src/worksheet/index';
import { makeBarChart, makeBarSeries, makeChartSpace, makeLineChart } from '../../../src/chart/chart';
import { makeOneCellAnchor } from '../../../src/drawing/anchor';
import { makeChartDrawingItem, makeDrawing, makePictureDrawingItem } from '../../../src/drawing/drawing';
import { loadImage } from '../../../src/drawing/image';
import { writeWorkbook } from '../_helpers';

const TINY_BLUE_PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAFklEQVR4nGP8z8DAwMDAxMDA8J+BAQAOAQHv6sTncgAAAABJRU5ErkJggg==';

describe('e2e 24 — multiple drawings on one sheet', () => {
  it('writes 24-multi-drawing.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Combo');

    // Source data
    setCell(ws, 1, 1, 'Quarter');
    setCell(ws, 1, 2, 'Sales');
    setCell(ws, 1, 3, 'Trend');
    const labels = ['Q1', 'Q2', 'Q3', 'Q4'];
    const sales = [120, 135, 158, 172];
    const trend = [100, 130, 160, 200];
    labels.forEach((q, i) => {
      setCell(ws, i + 2, 1, q);
      setCell(ws, i + 2, 2, sales[i] ?? 0);
      setCell(ws, i + 2, 3, trend[i] ?? 0);
    });

    const cat = { ref: 'Combo!$A$2:$A$5', cacheKind: 'str' as const, cache: labels };

    const barChart = makeChartSpace({
      title: 'Quarterly Sales',
      plotArea: {
        chart: makeBarChart({
          barDir: 'col',
          grouping: 'clustered',
          series: [
            makeBarSeries({
              idx: 0,
              val: { ref: 'Combo!$B$2:$B$5', cache: sales },
              cat,
              tx: { kind: 'ref', ref: 'Combo!$B$1' },
            }),
          ],
        }),
      },
      legend: { position: 'b' },
    });

    const lineChart = makeChartSpace({
      title: 'Trend',
      plotArea: {
        chart: makeLineChart({
          series: [
            makeBarSeries({
              idx: 0,
              val: { ref: 'Combo!$C$2:$C$5', cache: trend },
              cat,
              tx: { kind: 'ref', ref: 'Combo!$C$1' },
            }),
          ],
        }),
      },
      legend: { position: 'b' },
    });

    const pngBytes = Uint8Array.from(Buffer.from(TINY_BLUE_PNG_B64, 'base64'));
    const image = loadImage(pngBytes);

    ws.drawing = makeDrawing([
      makeChartDrawingItem(makeOneCellAnchor({ from: 'E2', widthPx: 480, heightPx: 320 }), { space: barChart }),
      makeChartDrawingItem(makeOneCellAnchor({ from: 'E20', widthPx: 480, heightPx: 320 }), { space: lineChart }),
      makePictureDrawingItem(makeOneCellAnchor({ from: 'N2', widthPx: 96, heightPx: 96 }), image),
    ]);

    const result = await writeWorkbook('24-multi-drawing.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
