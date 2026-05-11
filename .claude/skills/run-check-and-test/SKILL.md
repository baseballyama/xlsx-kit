---
name: run-check-and-test
description: Use after making code changes, before committing, before opening a PR, or whenever the user says "check it", "run tests", "verify", or "is this ready". Runs the project's quality gates — format, lint, types, tests, build — and reports failures.
---

# Quality gate

This skill runs the project's full check set and reports back. It is **language-
and tooling-agnostic**: read the project's manifest (`package.json`,
`Cargo.toml`, `pyproject.toml`, `Makefile`, etc.) to discover the actual scripts.
Do not assume.

## Discover the available scripts

Look in this order, depending on what's in the repo root:

| File                     | Where to look                                     |
| ------------------------ | ------------------------------------------------- |
| `package.json`           | `scripts` field (npm/pnpm/yarn/bun)               |
| `Cargo.toml`             | `cargo fmt`, `cargo clippy`, `cargo test`         |
| `pyproject.toml`         | `[tool.*]` sections, common: `ruff`, `pytest`, `mypy` |
| `Makefile`               | Common targets: `make check`, `make test`, `make lint` |
| `justfile` / `Taskfile`  | The recipes within                                |
| `mise.toml` / `.mise.toml` | Tasks defined under `[tasks]`                   |
| `CLAUDE.md`              | Project-specific overrides may be documented here |

If the user has explicitly told you the commands once before in the session, use
those. Otherwise re-discover — scripts get renamed.

## The gates, in order

Run in this order; stop and report on the first failure unless the user said "run
everything":

1. **Format** — does the code conform to the formatter?
   - Common: `prettier --check` / `cargo fmt --check` / `ruff format --check` / `gofmt -l`
   - If broken: try the auto-fix (`prettier --write` / `cargo fmt` / `ruff format`)
     before reporting. Re-run the check after fixing.

2. **Lint** — static analysis.
   - Common: `eslint`, `oxlint`, `clippy`, `ruff check`, `golangci-lint`
   - Some lint errors auto-fix; try `--fix` or equivalent before reporting only the
     unfixable ones.

3. **Type check** — for typed languages.
   - Common: `tsc --noEmit`, `cargo check`, `mypy`, `pyright`, `go vet`
   - Type errors are not auto-fixable; report them with file:line.

4. **Unit tests**
   - Common: `pnpm test`, `cargo test`, `pytest`, `go test ./...`
   - Run the full unit suite. If the user has already done this in the session and
     no test files have changed, you can skip and say so.

5. **Build** — for libraries that ship artifacts.
   - Common: `pnpm build`, `cargo build --release`, `python -m build`, `go build`
   - Catches errors in publish-time configuration that runtime tests miss.

6. **Project-specific gates** — packages may have additional gates worth running:
   - **Bundle size** (if the project is a library): `pnpm size`, `bundlewatch`,
     `size-limit`
   - **Knip / unused exports**: `pnpm knip`
   - **Security scanners**: `npm audit`, `cargo audit`, `pip-audit`
   - **Schema / contract checks**: project-defined

   Look in CI config (`.github/workflows/*.yml`) for what the maintainers consider
   important. If CI runs it, run it locally before pushing.

## Filtering to changed code

For monorepos and large codebases, running the full suite per change is slow. Most
ecosystems have a "changed files only" mode:

- **pnpm**: `pnpm --filter "...[origin/main]" <script>`
- **turbo**: `turbo run <script> --filter=[origin/main...]`
- **Cargo workspaces**: `cargo test -p <crate>` for the affected crate
- **pytest**: `pytest <changed-test-files>` for impact-scoped runs

Decision rule:

- Working interactively → filter to changed.
- Right before opening a PR → run the **full** suite, since CI will too.

## Reporting

For each gate, report one of:

- ✅ Passed
- ⚠️ Auto-fixed (state what was fixed)
- ❌ Failed (state which files / what the error was)

Don't dump the full tool output unless the user asks. Summarize:

```
Format:     ✅
Lint:       ⚠️  Auto-fixed 3 issues in src/foo.ts
Types:      ❌ src/bar.ts:42 — Property 'x' does not exist on type 'Y'
Tests:      (skipped — types failed)
Build:      (skipped — types failed)
```

If a gate fails, **stop and ask** before continuing the rest. The user may want to
fix the failure before running further gates.

## When the project has no obvious gate

If the repo doesn't define scripts (early-stage project, hand-written README), tell
the user that and offer to set them up — don't invent commands. Common minimum:

- A formatter (Prettier / `cargo fmt` / `ruff format` / `gofmt`)
- A linter (ESLint / Clippy / Ruff / golangci-lint)
- A test runner (Vitest / Jest / `cargo test` / pytest / `go test`)

Offer to wire these up via the `update-config` skill or by editing the manifest
directly, but get the user's go-ahead first.

## What this skill is NOT for

- **Long-running CI checks** (multi-minute integration tests, end-to-end browser
  tests, deployment dry-runs). Those belong on CI, not on the local pre-commit
  loop. If the project has both fast and slow gates, run only the fast set here,
  and tell the user the slow set will run on CI.
- **Coverage thresholds** as a gate. Coverage is informational, not blocking, in
  most projects. Only enforce it if the project's CI does.
- **Auto-fixing typos in test names or other "soft" issues** — those need human
  judgment.
