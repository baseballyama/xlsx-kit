// Set a formula. Optionally cache its evaluated value so Excel renders
// the result before recalculating on open.

import { setFormula } from 'ooxml-js/xlsx/cell';
import { saveWorkbook } from 'ooxml-js/xlsx/io';
import { toFile } from 'ooxml-js/node';
import { addWorksheet, createWorkbook } from 'ooxml-js/xlsx/workbook';
import { setCell } from 'ooxml-js/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Sheet1');

setCell(ws, 1, 1, 12);
setCell(ws, 2, 1, 18);
setCell(ws, 3, 1, 30);

setFormula(setCell(ws, 4, 1), 'SUM(A1:A3)', { cachedValue: 60 });

await saveWorkbook(wb, toFile('with-formulas.xlsx'));
