import { defineConfig } from 'vitest/config';

// Node-hosted suite. Browser-target tests will live under tests/browser/ and
// run via @vitest/browser in a later bootstrap commit.
//
// Vitest is the unified runner for both unit and integration suites; coverage
// is V8.

export default defineConfig({
  test: {
    environment: 'node',
    include: ['tests/**/*.test.ts'],
    exclude: ['tests/browser/**', 'tests/perf/**', 'tests/e2e/**', 'node_modules', 'dist', 'reference'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html'],
      include: ['src/**/*.ts'],
      exclude: ['src/**/*.d.ts', 'src/**/index.ts'],
    },
  },
});
