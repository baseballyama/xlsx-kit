// Make a cell clickable. The text is whatever you set on the cell;
// hyperlink wires up the URL underneath.

import { saveWorkbook, toFile } from 'openxml-js/node';
import { addWorksheet, createWorkbook } from 'openxml-js/workbook';
import { addUrlHyperlink, setCell } from 'openxml-js/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Links');

setCell(ws, 1, 1, 'Project home');
addUrlHyperlink(ws, 'A1', 'https://github.com/baseballyama/openxml-js', {
  tooltip: 'View on GitHub',
});

await saveWorkbook(wb, toFile('with-links.xlsx'));
