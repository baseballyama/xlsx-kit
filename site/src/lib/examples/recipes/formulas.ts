// Set a formula. Optionally cache its evaluated value so Excel renders
// the result before recalculating on open.

import { saveWorkbook, toFile } from 'openxml-js/node';
import { addWorksheet, createWorkbook } from 'openxml-js/workbook';
import { setCell, setCellFormula } from 'openxml-js/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Sheet1');

setCell(ws, 1, 1, 12);
setCell(ws, 2, 1, 18);
setCell(ws, 3, 1, 30);

setCellFormula(ws, 4, 1, 'SUM(A1:A3)', { cachedValue: 60 });

await saveWorkbook(wb, toFile('with-formulas.xlsx'));
