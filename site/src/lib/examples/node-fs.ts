// One-shot read + save direct from / to disk via the ooxml-js/node
// helpers, no manual fs glue needed.

import { loadWorkbook, saveWorkbook } from 'ooxml-js/xlsx/io';
import { fromFile, toFile } from 'ooxml-js/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
// ...mutate wb...
await saveWorkbook(wb, toFile('output.xlsx'));
