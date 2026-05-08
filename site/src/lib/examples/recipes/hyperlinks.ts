// Make a cell clickable. The text is whatever you set on the cell;
// hyperlink wires up the URL underneath.

import { saveWorkbook } from 'xlsx-craft/io';
import { toFile } from 'xlsx-craft/node';
import { addWorksheet, createWorkbook } from 'xlsx-craft/workbook';
import { setCell, setHyperlink } from 'xlsx-craft/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Links');

setCell(ws, 1, 1, 'Project home');
setHyperlink(ws, 'A1', {
  target: 'https://github.com/baseballyama/xlsx-craft',
  tooltip: 'View on GitHub',
});

await saveWorkbook(wb, toFile('with-links.xlsx'));
