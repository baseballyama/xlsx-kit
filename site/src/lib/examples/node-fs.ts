// One-shot read + save direct from / to disk via the xlsxify/node
// helpers, no manual fs glue needed.

import { loadWorkbook, saveWorkbook } from 'xlsxify/io';
import { fromFile, toFile } from 'xlsxify/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
// ...mutate wb...
await saveWorkbook(wb, toFile('output.xlsx'));
