// One-shot read + save direct from / to disk via the openxml-js/node
// helpers, no manual fs glue needed.

import { fromFile, loadWorkbook, saveWorkbook, toFile } from 'openxml-js/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
// ...mutate wb...
await saveWorkbook(wb, toFile('output.xlsx'));
