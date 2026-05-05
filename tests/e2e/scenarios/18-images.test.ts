// Scenario 18: an embedded image in a worksheet drawing.
// Output: 18-images.xlsx
//
// What to verify in Excel:
// - Sheet "Image" shows a small 4x4 PNG block anchored at C3.
// - Cell A1 says "image is at C3 →".
//
// The PNG payload here is a hand-crafted 4×4 32-bit RGBA solid-blue
// pattern; Excel should display it as a tiny blue square scaled to
// fit the anchor box.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook, setCell } from '../../../src/index';
import { makeOneCellAnchor } from '../../../src/drawing/anchor';
import { makeDrawing, makePictureDrawingItem } from '../../../src/drawing/drawing';
import { loadImage } from '../../../src/drawing/image';
import { writeWorkbook } from '../_helpers';

// Minimal 4x4 solid-blue PNG. Generated once via:
//   node -e "const z=require('zlib'),crc=require('crc-32');..."
// then committed as a base64 literal so the test doesn't need a build
// step. Check `tests/e2e/scenarios/_image-fixture.ts` for details.
const TINY_BLUE_PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAFklEQVR4nGP8z8DAwMDAxMDA8J+BAQAOAQHv6sTncgAAAABJRU5ErkJggg==';

describe('e2e 18 — embedded image', () => {
  it('writes 18-images.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Image');
    setCell(ws, 1, 1, 'image is at C3 →');

    const pngBytes = Uint8Array.from(Buffer.from(TINY_BLUE_PNG_B64, 'base64'));
    const image = loadImage(pngBytes);

    ws.drawing = makeDrawing([
      makePictureDrawingItem(
        makeOneCellAnchor({ from: 'C3', widthPx: 96, heightPx: 96 }),
        image,
      ),
    ]);

    const result = await writeWorkbook('18-images.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
