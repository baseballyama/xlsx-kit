// Heap-budget metric for the write-only path. docs/plan/06-streaming.md
// §3.4 sets an aspirational target of 100M cells in 1GB heap. The
// current `createWriteOnlyWorkbook` keeps each Cell as an object on the
// Worksheet rows Map until close(), so it can't meet that target
// end-to-end — but tracking the cells/MB ratio at modest cell counts
// surfaces regressions early.
//
// Excluded from the default `pnpm test` run (see vitest.config.ts).
// Run explicitly:
//   pnpm test:perf
//   PERF_HEAP_GATE=1 pnpm test:perf  # asserts the cells/MB floor
//
// Numbers depend on Node version + V8 GC mood. The default gate is off;
// the test is primarily a long-term tracking metric.

import { describe, expect, it } from 'vitest';
import { toBuffer } from '../../src/io/node';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only';

const ROWS = 100_000;
const COLS = 30;
const TOTAL_CELLS = ROWS * COLS;

const PERF_HEAP_GATE = process.env['PERF_HEAP_GATE'] === '1';
// Tracking floor for the *streaming-deflate* implementation (each row
// pushes through the deflate stream as ~64 KB chunks, no rowChunks /
// no Cell / no Map retention). Empirically lands around 88_000 cells
// per heapUsed MB on M-series Node 22 — set the gate at 50_000 to
// leave breathing room while catching real regressions. This puts us
// within striking distance of the docs target (100M cells in 1 GB
// heap → 100k cells/MB).
const FLOOR_CELLS_PER_HEAP_MB = 50_000;

describe('phase-4 perf — write-only heap budget', () => {
  it(
    `writes ${TOTAL_CELLS.toLocaleString()} cells and reports peak heap`,
    async () => {
      // Stabilise the baseline by GC-ing before measurement when --expose-gc
      // is active. Without it we just snapshot whatever V8 has resident.
      const gc = (globalThis as { gc?: () => void }).gc;
      if (typeof gc === 'function') gc();
      const before = process.memoryUsage();

      const sink = toBuffer();
      const wb = await createWriteOnlyWorkbook(sink);
      const ws = await wb.addWorksheet('Sheet1');
      let peakHeap = before.heapUsed;
      let peakRss = before.rss;

      const sampleEvery = ROWS / 50; // 50 samples — enough to spot the peak
      for (let r = 0; r < ROWS; r++) {
        const row = new Array<number>(COLS);
        for (let c = 0; c < COLS; c++) row[c] = r * COLS + c;
        await ws.appendRow(row);
        if (r % sampleEvery === 0) {
          const m = process.memoryUsage();
          if (m.heapUsed > peakHeap) peakHeap = m.heapUsed;
          if (m.rss > peakRss) peakRss = m.rss;
        }
      }
      await ws.close();
      const m = process.memoryUsage();
      if (m.heapUsed > peakHeap) peakHeap = m.heapUsed;
      if (m.rss > peakRss) peakRss = m.rss;
      await wb.finalize();
      const archiveBytes = sink.result().byteLength;

      const heapMb = (peakHeap - before.heapUsed) / (1024 * 1024);
      const rssMb = (peakRss - before.rss) / (1024 * 1024);
      const cellsPerHeapMb = Math.round(TOTAL_CELLS / Math.max(heapMb, 1));

      process.stderr.write(
        `[perf-heap] ${TOTAL_CELLS.toLocaleString()} cells → peak heap +${heapMb.toFixed(1)} MB, peak rss +${rssMb.toFixed(1)} MB, ${cellsPerHeapMb.toLocaleString()} cells/heap-MB; archive ${archiveBytes.toLocaleString()} bytes\n`,
      );

      expect(archiveBytes).toBeGreaterThan(0);
      if (PERF_HEAP_GATE) {
        expect(cellsPerHeapMb).toBeGreaterThanOrEqual(FLOOR_CELLS_PER_HEAP_MB);
      }
    },
    /* timeout */ 10 * 60_000,
  );
});
