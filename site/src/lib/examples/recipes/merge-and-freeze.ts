// Merge a header range across the top row and freeze the first row so
// it stays visible while scrolling.

import { saveWorkbook } from 'xlsx-craft/io';
import { toFile } from 'xlsx-craft/node';
import { centerCell, setBold } from 'xlsx-craft/styles';
import { addWorksheet, createWorkbook } from 'xlsx-craft/workbook';
import {
  makeFreezePane,
  makeSheetView,
  mergeCells,
  setCell,
} from 'xlsx-craft/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Report');

const title = setCell(ws, 1, 1, 'Q2 financial summary');
setBold(wb, title);
centerCell(wb, title);
mergeCells(ws, 'A1:E1');

ws.views.push(makeSheetView({ pane: makeFreezePane('A2') }));

await saveWorkbook(wb, toFile('merged-frozen.xlsx'));
