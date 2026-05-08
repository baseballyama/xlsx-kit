// Add a clustered column chart driven by a data range on the same sheet.

import { makeBarChart, makeBarSeries, makeChartSpace } from 'openxml-js/chart';
import { addChartAt } from 'openxml-js/drawing';
import { saveWorkbook, toFile } from 'openxml-js/node';
import { addWorksheet, createWorkbook } from 'openxml-js/workbook';
import { setCell } from 'openxml-js/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Sales');

setCell(ws, 1, 1, 'Region');
setCell(ws, 1, 2, 'Revenue');
setCell(ws, 2, 1, 'NA');
setCell(ws, 2, 2, 12_400);
setCell(ws, 3, 1, 'EU');
setCell(ws, 3, 2, 9_800);
setCell(ws, 4, 1, 'APAC');
setCell(ws, 4, 2, 7_300);

const chart = makeBarChart({
  barDir: 'col',
  grouping: 'clustered',
  series: [
    makeBarSeries({
      idx: 0,
      tx: { kind: 'literal', value: 'Revenue' },
      cat: { ref: 'Sales!$A$2:$A$4' },
      val: { ref: 'Sales!$B$2:$B$4' },
    }),
  ],
});

const space = makeChartSpace({
  plotArea: { chart },
  title: 'Revenue by region',
  legend: { position: 'r' },
});

addChartAt(ws, 'D2', { space }, { widthPx: 480, heightPx: 320 });

await saveWorkbook(wb, toFile('chart.xlsx'));
