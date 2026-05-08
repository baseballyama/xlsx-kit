// Open a workbook and walk every cell on the first sheet.

import { loadWorkbook } from 'xlsx-craft/io';
import { fromFile } from 'xlsx-craft/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
const first = wb.sheets[0];
if (first?.kind === 'worksheet') {
  for (const row of first.sheet.rows.values()) {
    for (const cell of row.values()) {
      console.log(`${cell.row},${cell.col}: ${String(cell.value)}`);
    }
  }
}
