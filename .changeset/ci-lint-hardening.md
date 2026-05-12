---
'xlsx-kit': patch
---

chore(ci): tighten the CI gate and surface CLAUDE.md hard rules in the linter.

- The test matrix now runs on macOS and Windows alongside Ubuntu. Path-separator and TextDecoder differences that ubuntu-only CI would silently miss are now exercised on every PR. Each OS installs `libxml2-utils` / `libxml2` so the ECMA-376 conformance gate never silently degrades to "no schema validation ran".

- A dedicated `perf` job runs `pnpm test:perf` with `PERF_GATE=1`, promoting the throughput / heap thresholds in `tests/perf/` from informational to fatal. CI runners are noisier than the M1 baseline the gates target — if a specific Node minor turns flaky, retune the thresholds in `vitest.perf.config.ts` rather than reverting this gate.

- `typescript/no-explicit-any` is `error` in `src/` and `off` in `tests/` (where `as any` is a legitimate way to exercise error paths). The remaining intentional `any` in `src/schema/core.ts` is annotated with `// oxlint-disable-next-line typescript/no-explicit-any` and an explanation comment.
