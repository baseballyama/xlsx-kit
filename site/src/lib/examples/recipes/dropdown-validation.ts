// Add a list-type data validation — gives the user a dropdown of
// allowed values when they click into the range.

import { saveWorkbook } from 'xlsxify/io';
import { toFile } from 'xlsxify/node';
import { addWorksheet, createWorkbook } from 'xlsxify/workbook';
import { addDataValidation, makeDataValidation, setCell } from 'xlsxify/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Form');

setCell(ws, 1, 1, 'Status');
addDataValidation(
  ws,
  makeDataValidation({
    type: 'list',
    sqref: 'B1:B100',
    formula1: '"Open,In progress,Closed"',
    prompt: 'Pick a status',
    errorTitle: 'Invalid value',
    error: 'Pick one of the listed values.',
  }),
);

await saveWorkbook(wb, toFile('with-dropdown.xlsx'));
