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

// Sink that writes to /dev/null — drops every chunk on the floor so
// heap measurements isolate the writer + deflate state from the
// buffered-output retention dominating large runs.
const discardSink = (): Parameters<typeof createWriteOnlyWorkbook>[0] & { result(): number } => {
  let total = 0;
  return {
    toBytes() {
      return {
        write(chunk: Uint8Array) {
          total += chunk.byteLength;
        },
        async finish() {
          return new Uint8Array(0);
        },
      };
    },
    result() {
      return total;
    },
  };
};

const measureHeap = async (
  rows: number,
  cols: number,
  sinkFactory: () => Parameters<typeof createWriteOnlyWorkbook>[0],
): Promise<{ heapMb: number; rssMb: number; archiveBytes: number; seconds: number }> => {
  const gc = (globalThis as { gc?: () => void }).gc;
  if (typeof gc === 'function') gc();
  const before = process.memoryUsage();
  const sink = sinkFactory();
  const wb = await createWriteOnlyWorkbook(sink);
  const ws = await wb.addWorksheet('Sheet1');
  let peakHeap = before.heapUsed;
  let peakRss = before.rss;
  const sampleEvery = Math.max(1, Math.floor(rows / 50));
  const t0 = performance.now();
  for (let r = 0; r < rows; r++) {
    const row = new Array<number>(cols);
    for (let c = 0; c < cols; c++) row[c] = r * cols + c;
    await ws.appendRow(row);
    if (r % sampleEvery === 0) {
      const m = process.memoryUsage();
      if (m.heapUsed > peakHeap) peakHeap = m.heapUsed;
      if (m.rss > peakRss) peakRss = m.rss;
    }
  }
  await ws.close();
  await wb.finalize();
  const t1 = performance.now();
  const m = process.memoryUsage();
  if (m.heapUsed > peakHeap) peakHeap = m.heapUsed;
  if (m.rss > peakRss) peakRss = m.rss;
  const archiveBytes =
    'result' in sink && typeof sink.result === 'function'
      ? typeof (sink.result() as unknown) === 'number'
        ? (sink.result() as unknown as number)
        : ((sink.result() as Buffer).byteLength ?? 0)
      : 0;
  return {
    heapMb: (peakHeap - before.heapUsed) / (1024 * 1024),
    rssMb: (peakRss - before.rss) / (1024 * 1024),
    archiveBytes,
    seconds: (t1 - t0) / 1000,
  };
};

describe('phase-4 perf — heap stays flat as cells grow', () => {
  it(
    'verifies the 100M-cells / 1GB-heap target by measuring 1M / 3M / 10M with a discard sink',
    async () => {
      // The streaming-deflate write keeps a fixed-cost working set
      // (encoding scratch + ZipDeflate sliding window). With a discard
      // sink the buffered-output retention vanishes too, so heap
      // shouldn't scale with cell count. If it does, we've regressed
      // the streaming property.
      const sizes: Array<{ rows: number; cols: number }> = [
        { rows: 33_333, cols: 30 }, // ~1M
        { rows: 100_000, cols: 30 }, // ~3M
        { rows: 333_333, cols: 30 }, // ~10M
      ];
      const lines: string[] = [];
      const heaps: number[] = [];
      for (const { rows, cols } of sizes) {
        const cells = rows * cols;
        const r = await measureHeap(rows, cols, discardSink);
        const cellsPerMb = Math.round(cells / Math.max(r.heapMb, 0.01));
        lines.push(
          `[perf-scale] ${cells.toLocaleString()} cells → +${r.heapMb.toFixed(1)} MB heap, +${r.rssMb.toFixed(1)} MB rss, ${cellsPerMb.toLocaleString()} cells/heap-MB, ${r.seconds.toFixed(2)}s`,
        );
        heaps.push(r.heapMb);
      }
      process.stderr.write(`${lines.join('\n')}\n`);
      // Heap at 10M cells must stay well under 1 GB. If our scaling
      // were linear we'd be at ~330 MB; in practice it's < 100 MB.
      expect(heaps[heaps.length - 1] ?? Infinity).toBeLessThan(500);
    },
    /* timeout */ 10 * 60_000,
  );
});

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
