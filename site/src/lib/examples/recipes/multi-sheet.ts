// Build several worksheets in one workbook and use named ranges
// to refer between them.

import { setFormula } from 'xlsx-craft/cell';
import { saveWorkbook } from 'xlsx-craft/io';
import { toFile } from 'xlsx-craft/node';
import { addDefinedName, addWorksheet, createWorkbook } from 'xlsx-craft/workbook';
import { setCell } from 'xlsx-craft/worksheet';

const wb = createWorkbook();
const inputs = addWorksheet(wb, 'Inputs');
const summary = addWorksheet(wb, 'Summary');

setCell(inputs, 1, 1, 'Revenue');
setCell(inputs, 1, 2, 100_000);
setCell(inputs, 2, 1, 'Cost');
setCell(inputs, 2, 2, 65_000);

addDefinedName(wb, { name: 'Revenue', value: 'Inputs!$B$1' });
addDefinedName(wb, { name: 'Cost', value: 'Inputs!$B$2' });

setCell(summary, 1, 1, 'Margin');
setFormula(setCell(summary, 1, 2), '(Revenue - Cost) / Revenue', { cachedValue: 0.35 });

await saveWorkbook(wb, toFile('multi-sheet.xlsx'));
