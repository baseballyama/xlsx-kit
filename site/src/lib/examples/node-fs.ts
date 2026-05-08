// One-shot read + save direct from / to disk via the xlsxlite/node
// helpers, no manual fs glue needed.

import { loadWorkbook, saveWorkbook } from 'xlsxlite/io';
import { fromFile, toFile } from 'xlsxlite/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
// ...mutate wb...
await saveWorkbook(wb, toFile('output.xlsx'));
