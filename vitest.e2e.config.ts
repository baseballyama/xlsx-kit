import { defineConfig } from 'vitest/config';

// E2E config — runs `tests/e2e/scenarios/*.test.ts` which generate
// real .xlsx / .xlsm files into `tests/e2e/output/`. Each scenario
// emits one file; `pnpm test:e2e` builds them all so the user can
// open the directory in Excel / LibreOffice / Google Sheets and
// visually verify the output matches the README's checklist.
//
// The default `pnpm test` config (`vitest.config.ts`) already excludes
// tests/e2e/** so this config is only used by the explicit script.

export default defineConfig({
  test: {
    environment: 'node',
    include: ['tests/e2e/scenarios/**/*.test.ts'],
    exclude: ['node_modules', 'dist', 'reference'],
    // Each scenario writes one file — sequential keeps the [e2e]
    // log output coherent. (Generation is fast; parallelism saves
    // very little.)
    fileParallelism: false,
  },
});
