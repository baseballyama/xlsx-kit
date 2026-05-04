import { defineConfig } from 'tsup';

// Bootstrap-stage tsup config: a single ESM entry from src/index.
//
// The eventual matrix (per docs/plan/11-build-publish.md §1.3) builds many
// subpath entries (read / write / streaming / styles / chart / chart-extended
// / drawing / pivot / schema / io-node / io-browser) with node and browser
// platform variants. That config lands when we have actual subpackage entries
// to point it at — not before. Keeping the surface narrow now means a CI
// build smoke can pass with no dead exports.

export default defineConfig({
  entry: { index: 'src/index.ts' },
  format: ['esm'],
  target: 'es2022',
  platform: 'neutral',
  sourcemap: true,
  clean: true,
  treeshake: true,
  dts: false,
  splitting: false,
  minify: false,
  outDir: 'dist',
  // Force `.mjs` so the package.json exports map (which points at `.mjs`)
  // resolves regardless of `type: module` defaulting to `.js`.
  outExtension: () => ({ js: '.mjs' }),
});
