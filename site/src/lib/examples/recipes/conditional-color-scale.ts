// Color-scale rule: red for low values, yellow for the middle, green
// for high. Excel's classic 3-color heat-map.

import { saveWorkbook, toFile } from 'openxml-js/node';
import { addWorksheet, createWorkbook } from 'openxml-js/workbook';
import { addColorScaleRule, setCell } from 'openxml-js/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Heat');

for (let r = 1; r <= 10; r++) setCell(ws, r, 1, Math.round(Math.random() * 100));

addColorScaleRule(ws, 'A1:A10', {
  cfvos: [
    { type: 'min' },
    { type: 'percentile', val: '50' },
    { type: 'max' },
  ],
  colors: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
});

await saveWorkbook(wb, toFile('heatmap.xlsx'));
