// Scenario 14: multi-sheet workbook with visible / hidden / veryHidden
// states and a chartsheet for variety.
// Output: 14-multi-sheet.xlsx
//
// What to verify in Excel:
// - Tabs visible: "Visible1", "Visible2", and the chartsheet "Chart Tab".
// - "Hidden" sheet not in tab list; right-click any tab → Unhide → see it.
// - "VeryHidden" not even in the Unhide dialog (only visible via VBA / our API).
// - The chartsheet renders a full-screen pie chart of {A: 30, B: 50, C: 20}.

import { describe, expect, it } from 'vitest';
import { addChartsheet, addWorksheet, createWorkbook } from '../../../src/xlsx/workbook/index';
import { setCell } from '../../../src/xlsx/worksheet/index';
import { makeChartSpace, makePieChart, makeBarSeries } from '../../../src/xlsx/chart/chart';
import { writeWorkbook } from '../_helpers';

describe('e2e 14 — multi-sheet (visible / hidden / veryHidden / chartsheet)', () => {
  it('writes 14-multi-sheet.xlsx', async () => {
    const wb = createWorkbook();

    const v1 = addWorksheet(wb, 'Visible1');
    const labels = ['A', 'B', 'C'];
    const values = [30, 50, 20];
    setCell(v1, 1, 1, 'Group');
    setCell(v1, 1, 2, 'Value');
    labels.forEach((lab, i) => {
      setCell(v1, i + 2, 1, lab);
      setCell(v1, i + 2, 2, values[i] ?? 0);
    });

    const v2 = addWorksheet(wb, 'Visible2');
    setCell(v2, 1, 1, 'visible-2');

    const hidden = addWorksheet(wb, 'Hidden', { state: 'hidden' });
    setCell(hidden, 1, 1, 'hidden — right-click → Unhide');

    const vh = addWorksheet(wb, 'VeryHidden', { state: 'veryHidden' });
    setCell(vh, 1, 1, 'veryHidden — only via VBA / our API');

    const pie = makePieChart({
      series: [
        makeBarSeries({
          idx: 0,
          val: { ref: 'Visible1!$B$2:$B$4', cache: values },
          cat: { ref: 'Visible1!$A$2:$A$4', cacheKind: 'str', cache: labels },
        }),
      ],
    });
    addChartsheet(wb, 'Chart Tab', {
      chart: {
        space: makeChartSpace({
          title: 'Distribution',
          plotArea: { chart: pie },
          legend: { position: 'r' },
        }),
      },
    });

    const result = await writeWorkbook('14-multi-sheet.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
