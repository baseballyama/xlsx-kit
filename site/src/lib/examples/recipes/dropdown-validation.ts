// Add a list-type data validation — gives the user a dropdown of
// allowed values when they click into the range.

import { saveWorkbook, toFile } from 'openxml-js/node';
import { addWorksheet, createWorkbook } from 'openxml-js/workbook';
import { addListValidation, setCell } from 'openxml-js/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Form');

setCell(ws, 1, 1, 'Status');
addListValidation(ws, 'B1:B100', ['Open', 'In progress', 'Closed'], {
  prompt: 'Pick a status',
  errorTitle: 'Invalid value',
  error: 'Pick one of the listed values.',
});

await saveWorkbook(wb, toFile('with-dropdown.xlsx'));
