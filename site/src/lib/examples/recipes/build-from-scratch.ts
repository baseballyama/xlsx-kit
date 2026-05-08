// Build a one-sheet workbook from scratch and write it to disk.

import { saveWorkbook } from 'xlsx-kit/io';
import { toFile } from 'xlsx-kit/node';
import { addWorksheet, createWorkbook } from 'xlsx-kit/workbook';
import { setCell } from 'xlsx-kit/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Quarterly');

setCell(ws, 1, 1, 'Quarter');
setCell(ws, 1, 2, 'Revenue');
setCell(ws, 2, 1, 'Q1');
setCell(ws, 2, 2, 12_400);
setCell(ws, 3, 1, 'Q2');
setCell(ws, 3, 2, 15_900);

await saveWorkbook(wb, toFile('quarterly.xlsx'));
