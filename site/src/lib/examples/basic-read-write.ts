// Read an xlsx, mutate one cell, write it back.
//
// This file is imported as ?raw into the docs site so the snippet shown to
// readers is exactly what svelte-check / tsc compiled — if an API rename
// breaks this import, the docs build fails before deploy.

import { loadWorkbook, workbookToBytes } from 'ooxml-js/xlsx/io';
import { fromBuffer } from 'ooxml-js/node';
import { setCell } from 'ooxml-js/xlsx/worksheet';
import { readFile, writeFile } from 'node:fs/promises';

const wb = await loadWorkbook(fromBuffer(await readFile('input.xlsx')));
const ref = wb.sheets[0];
if (ref?.kind === 'worksheet') {
  setCell(ref.sheet, 1, 1, 'Hello from ooxml-js');
}
await writeFile('output.xlsx', await workbookToBytes(wb));
