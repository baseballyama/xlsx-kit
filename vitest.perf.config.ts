import { defineConfig } from 'vitest/config';

// Perf-only config. `pnpm test:perf` runs tests under tests/perf/ which the
// default config explicitly excludes (so `pnpm test` stays fast).
//
// The write-only throughput gate measures ≥500k cells/s for the 100k×30 shape
// on an M1 baseline. Set PERF_GATE=1 to fail the run on regressions.

export default defineConfig({
  test: {
    environment: 'node',
    include: ['tests/perf/**/*.test.ts'],
    // Bench files run via `pnpm bench` (vitest bench), not the test runner.
    exclude: ['tests/perf/**/*.bench.ts', 'node_modules', 'dist', 'reference'],
  },
});
