// Make a cell clickable. The text is whatever you set on the cell;
// hyperlink wires up the URL underneath.

import { saveWorkbook } from 'xlsxify/io';
import { toFile } from 'xlsxify/node';
import { addWorksheet, createWorkbook } from 'xlsxify/workbook';
import { setCell, setHyperlink } from 'xlsxify/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Links');

setCell(ws, 1, 1, 'Project home');
setHyperlink(ws, 'A1', {
  target: 'https://github.com/baseballyama/xlsxify',
  tooltip: 'View on GitHub',
});

await saveWorkbook(wb, toFile('with-links.xlsx'));
