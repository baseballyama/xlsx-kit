// One-shot read + save direct from / to disk via the xlsx-kit/node
// helpers, no manual fs glue needed.

import { loadWorkbook, saveWorkbook } from 'xlsx-kit/io';
import { fromFile, toFile } from 'xlsx-kit/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
// ...mutate wb...
await saveWorkbook(wb, toFile('output.xlsx'));
