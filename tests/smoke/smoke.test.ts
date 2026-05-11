import { describe, expect, it } from 'vitest';

// Phase 0 smoke test: confirms the runner, TypeScript build pipeline and
// public entry are wired up end-to-end. Replaced by real phase 1 suites
// once src/io/ lands.

describe('phase-0 smoke', () => {
  it('vitest can import a section subpath', async () => {
    const mod = await import('../../src/workbook/index');
    expect(mod).toBeTypeOf('object');
    expect(typeof mod.createWorkbook).toBe('function');
  });

  it('Node version supports the engines.node floor', () => {
    const [major] = process.versions.node.split('.').map(Number);
    expect(major).toBeGreaterThanOrEqual(18);
  });
});
