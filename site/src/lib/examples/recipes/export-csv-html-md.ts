// Render a worksheet as CSV / HTML / Markdown / plain text. The
// `getWorksheetAs*` helpers auto-detect the data extent, so you don't
// have to compute a range yourself.

import { loadWorkbook } from 'openxml-js/io';
import { fromFile } from 'openxml-js/node';
import {
  getWorksheetAsCsv,
  getWorksheetAsHtml,
  getWorksheetAsMarkdownTable,
  getWorksheetAsTextTable,
} from 'openxml-js/worksheet';
import { writeFile } from 'node:fs/promises';

const wb = await loadWorkbook(fromFile('input.xlsx'));
const sheet = wb.sheets[0];
if (sheet?.kind !== 'worksheet') throw new Error('first sheet is a chartsheet');
const ws = sheet.sheet;

await writeFile('out.csv', getWorksheetAsCsv(ws));
await writeFile('out.html', getWorksheetAsHtml(wb, ws));
await writeFile('out.md', getWorksheetAsMarkdownTable(ws));
await writeFile('out.txt', getWorksheetAsTextTable(ws));
