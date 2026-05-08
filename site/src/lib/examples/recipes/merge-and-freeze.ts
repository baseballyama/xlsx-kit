// Merge a header range across the top row and freeze the first row so
// it stays visible while scrolling.

import { saveWorkbook } from 'xlsxify/io';
import { toFile } from 'xlsxify/node';
import { centerCell, setBold } from 'xlsxify/styles';
import { addWorksheet, createWorkbook } from 'xlsxify/workbook';
import {
  makeFreezePane,
  makeSheetView,
  mergeCells,
  setCell,
} from 'xlsxify/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Report');

const title = setCell(ws, 1, 1, 'Q2 financial summary');
setBold(wb, title);
centerCell(wb, title);
mergeCells(ws, 'A1:E1');

ws.views.push(makeSheetView({ pane: makeFreezePane('A2') }));

await saveWorkbook(wb, toFile('merged-frozen.xlsx'));
