// Perf bench for the write-only streaming path. Acceptance: ≥500k cells/s on
// the M1 baseline for the 100k-row × 30-col shape (~3M cells). Vitest's `bench`
// API runs a sized N=2 batch (2 inner iterations × the warm-up + measurement
// passes vitest does on its own) and reports ops/sec — reciprocal gives seconds
// per archive, which we convert to cells/s in the throughput.test.ts gate
// alongside.
//
// Run with: pnpm bench Excluded from `pnpm test` via vitest.config.ts
// (`exclude: tests/perf/**`).

import { bench, describe } from 'vitest';
import { toBuffer } from '../../src/io/node';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only';

const ROWS = 100_000;
const COLS = 30;

const buildRow = (rowIdx: number): number[] => {
  const row = new Array<number>(COLS);
  for (let c = 0; c < COLS; c++) row[c] = rowIdx * COLS + c;
  return row;
};

describe('write-only — 100k rows × 30 cols (~3M cells)', () => {
  bench(
    'createWriteOnlyWorkbook → appendRow ×100k → finalize',
    async () => {
      const sink = toBuffer();
      const wb = await createWriteOnlyWorkbook(sink);
      const ws = await wb.addWorksheet('Sheet1');
      for (let r = 0; r < ROWS; r++) {
        await ws.appendRow(buildRow(r));
      }
      await ws.close();
      await wb.finalize();
    },
    { iterations: 2, warmupIterations: 1, warmupTime: 0 },
  );
});
