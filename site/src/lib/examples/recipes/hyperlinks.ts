// Make a cell clickable. The text is whatever you set on the cell;
// hyperlink wires up the URL underneath.

import { saveWorkbook } from 'ooxml-js/xlsx/io';
import { toFile } from 'ooxml-js/node';
import { addWorksheet, createWorkbook } from 'ooxml-js/xlsx/workbook';
import { setCell, setHyperlink } from 'ooxml-js/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Links');

setCell(ws, 1, 1, 'Project home');
setHyperlink(ws, 'A1', {
  target: 'https://github.com/baseballyama/ooxml-js',
  tooltip: 'View on GitHub',
});

await saveWorkbook(wb, toFile('with-links.xlsx'));
