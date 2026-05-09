// Throughput gate for the write-only path. Sits alongside the bench so CI can
// fail fast if a regression drops cells/s below the acceptance floor (≥500k
// cells/s on M1).
//
// Excluded from the default `pnpm test` run (see vitest.config.ts `exclude:
// ['tests/perf/**']`). Run explicitly:
//
//     pnpm test:perf       # uses tests/perf/* as the include pattern
//     PERF_GATE=1 pnpm test:perf  # also assert the throughput floor
//
// The bench uses a 100k-row × 30-col shape (~3M cells) — the same shape the
// bench file measures; smaller shapes don't amortise the finalize-time
// stylesheet/sst flush.

import { describe, expect, it } from 'vitest';
import { toBuffer } from '../../src/io/node';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only';

const ROWS = 100_000;
const COLS = 30;
const TOTAL_CELLS = ROWS * COLS;

// PERF_GATE off by default: laptops idle / CI under load can dip below the
// M1-baseline floor without the code being wrong. Set the env var when you want
// a hard assertion (release branches, perf-regression PRs).
const PERF_GATE = process.env['PERF_GATE'] === '1';
const FLOOR_CELLS_PER_SEC = 500_000;

const buildRow = (rowIdx: number): number[] => {
  const row = new Array<number>(COLS);
  for (let c = 0; c < COLS; c++) row[c] = rowIdx * COLS + c;
  return row;
};

const ITERATIONS = 3;

const measureOnce = async (): Promise<{ seconds: number; archiveBytes: number }> => {
  const sink = toBuffer();
  const wb = await createWriteOnlyWorkbook(sink);
  const ws = await wb.addWorksheet('Sheet1');
  const t0 = performance.now();
  for (let r = 0; r < ROWS; r++) {
    await ws.appendRow(buildRow(r));
  }
  await ws.close();
  await wb.finalize();
  const t1 = performance.now();
  return { seconds: (t1 - t0) / 1000, archiveBytes: sink.result().byteLength };
};

describe('phase-4 perf — write-only throughput', () => {
  it(
    `writes ${ROWS} × ${COLS} = ${TOTAL_CELLS.toLocaleString()} cells and reports cells/s`,
    async () => {
      // Best-of-N: shared CPUs / thermal throttling create wide variance in the
      // per-run number, but the best-case run reflects the pipeline's real
      // ceiling.
      const runs: Array<{ seconds: number; archiveBytes: number }> = [];
      for (let i = 0; i < ITERATIONS; i++) runs.push(await measureOnce());
      const bestSeconds = Math.min(...runs.map((r) => r.seconds));
      const bestCellsPerSec = Math.round(TOTAL_CELLS / bestSeconds);
      const archiveBytes = runs[0]?.archiveBytes ?? 0;

      const summaries = runs
        .map((r, i) => `#${i + 1} ${(TOTAL_CELLS / r.seconds).toFixed(0)} cells/s (${r.seconds.toFixed(2)}s)`)
        .join(' · ');
      process.stderr.write(
        `[perf] ${TOTAL_CELLS.toLocaleString()} cells × ${ITERATIONS} runs → best ${bestCellsPerSec.toLocaleString()} cells/s; archive ${archiveBytes.toLocaleString()} bytes\n        runs: ${summaries}\n`,
      );

      expect(archiveBytes).toBeGreaterThan(0);
      if (PERF_GATE) {
        expect(bestCellsPerSec).toBeGreaterThanOrEqual(FLOOR_CELLS_PER_SEC);
      }
    },
    /* timeout */ 10 * 60_000,
  );
});
