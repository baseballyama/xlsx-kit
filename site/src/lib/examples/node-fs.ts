// One-shot read + save direct from / to disk via the xlsx-craft/node
// helpers, no manual fs glue needed.

import { loadWorkbook, saveWorkbook } from 'xlsx-craft/io';
import { fromFile, toFile } from 'xlsx-craft/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
// ...mutate wb...
await saveWorkbook(wb, toFile('output.xlsx'));
