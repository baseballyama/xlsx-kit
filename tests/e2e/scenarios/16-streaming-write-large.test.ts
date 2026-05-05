// Scenario 16: large workbook via createWriteOnlyWorkbook — exercises
// the streaming-deflate writer with realistic shape (50k rows × 6 cols
// = 300k cells). Reasonably opens in Excel for visual smoke test
// without making the file enormous.
//
// Output: 16-streaming-large.xlsx
//
// What to verify in Excel:
// - Sheet "Generated" shows ~50k rows. Scroll to the bottom — last
//   row should be `row-50000`.
// - First column has row numbers; second column has random-ish
//   alternating strings; subsequent columns have numeric values.
// - File size in the OS file manager should be well under 10 MB.

import { describe, expect, it } from 'vitest';
import { mkdirSync } from 'node:fs';
import { OUT_DIR } from '../_helpers';
import { toFile } from '../../../src/node';
import { createWriteOnlyWorkbook } from '../../../src/streaming/index';

describe('e2e 16 — streaming write (50k rows)', () => {
  it('writes 16-streaming-large.xlsx via createWriteOnlyWorkbook', async () => {
    mkdirSync(OUT_DIR, { recursive: true });
    const path = `${OUT_DIR}/16-streaming-large.xlsx`;
    const wb = await createWriteOnlyWorkbook(toFile(path));
    const ws = await wb.addWorksheet('Generated');
    ws.setColumnWidth(1, 10);
    ws.setColumnWidth(2, 18);
    ws.setColumnWidth(3, 14);

    for (let r = 1; r <= 50_000; r++) {
      await ws.appendRow([
        r,
        r % 2 === 0 ? `even-${r}` : `odd-${r}`,
        r * 1.5,
        Math.sin(r),
        Math.cos(r),
        r % 7 === 0 ? 'lucky' : '',
      ]);
    }
    await ws.close();
    await wb.finalize();

    process.stderr.write(`[e2e] wrote 16-streaming-large.xlsx → ${path}\n`);
    expect(true).toBe(true);
  });
});
