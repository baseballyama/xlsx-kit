// Insert a PNG / JPEG image at a cell anchor. Format and dimensions
// are auto-detected from the bytes, so loadImage is the only call.

import { addImageAt, loadImage } from 'xlsx-craft/drawing';
import { saveWorkbook } from 'xlsx-craft/io';
import { toFile } from 'xlsx-craft/node';
import { addWorksheet, createWorkbook } from 'xlsx-craft/workbook';
import { readFile } from 'node:fs/promises';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Cover');

const image = loadImage(await readFile('logo.png'));
addImageAt(ws, 'B2', image, { widthPx: 200, heightPx: 80 });

await saveWorkbook(wb, toFile('with-image.xlsx'));
