---
name: full-code-review
description: Use before opening a PR to do a maintainer-perspective review of a branch. Walks through "why does this exist", "is the design right", "is the implementation correct", and "are tests / docs / changelog complete". Invoked when the user says "review my branch", "do a full review", "is this PR ready", or "/full-code-review".
---

# Full code review

Before opening a PR — or before merging your own — review the branch as if you were
a maintainer who didn't write it. The author's view is "what I built"; the
maintainer's view is "what this costs forever." This skill walks the second view.

The skill is **collaborative**: each phase produces a finding, you and the user
agree on what to do about it, and only then do you move to the next phase. Do not
batch all phases and dump the result at the end.

## Step 0 — Establish the baseline

Decide what you're comparing against and what the diff actually contains.

```bash
# Determine the base branch (in priority order):
# 1. User explicitly said "compare against X" → use X
# 2. PR exists → gh pr view <num> --json baseRefName
# 3. Default → origin/main (or the project's primary branch)

git fetch origin
git diff --stat ${BASE}...HEAD
git log --oneline ${BASE}...HEAD
git diff --name-only ${BASE}...HEAD
```

Categorize the changed files (use whatever buckets fit the project):

- Public API surface (entry points, type exports)
- Internal implementation
- Tests
- Documentation
- Build / CI / tooling
- Dependencies (lockfile, package manifests)

**Report to user**: which base branch you're using, the diff size, and the
categorization. If the diff is huge (thousands of lines or many unrelated buckets),
flag that this PR may need to be split — and confirm whether to continue.

## Phase 1 — Why does this change exist?

The hardest review question is **"should this even ship?"** Before judging the code,
make sure the change has a real reason to exist.

Pull what's available:

- PR description, linked issues, commit messages
- Any design discussion the user can paste in
- Connect to the `CLAUDE.md` policies — especially **"one way to do one thing"**

Confirm with the user:

- **Problem**: what is the user-facing problem this solves?
- **Audience**: who specifically is affected? (new users, advanced users, a specific
  integration?)
- **Alternatives**: was a smaller, doc-only, or "use the existing path" answer
  considered?
- **Scope fit**: does this stay inside this project's stated scope?

Possible verdicts:

| Verdict      | Meaning                                                                 | Next step                                          |
| ------------ | ----------------------------------------------------------------------- | -------------------------------------------------- |
| Justified    | Real problem, fits scope, no existing path covers it                    | Continue to Phase 2                                |
| Already covered | Existing API solves this; the PR is a parallel path → reject by policy | Stop. Recommend closing the PR with a code pointer. |
| Out of scope | Inside the user's universe, but not this library's scope                | Stop. Recommend redirecting to the right project.  |
| Unclear      | Reason isn't articulated                                                | Ask the user for the missing context              |

Don't proceed past Phase 1 until the verdict is "Justified."

## Phase 2 — Is the design right?

Now look at the **shape** of the change, not the code. Use the diff and a sketch of
the new public surface. Specifically check:

1. **API surface impact**
   - What new public exports / types / commands does this introduce?
   - What existing exports change behavior, signature, or semantics?
   - Are any breaking changes hidden as "fixes"? (Renames, type narrowings,
     defaults flipping, error types changing all count as breaking.)

2. **"One way to do one thing" check**
   - Does the new path duplicate a capability already in the API?
   - If yes: reject the design unless this is a *strictly better* replacement (and
     the old path is being removed in the same change).

3. **Consistency with existing patterns**
   - Naming, error types, async patterns, return shapes, file organization — does it
     look like the rest of the codebase, or did the author invent a one-off pattern?
   - If the pattern is intentionally new, is the reason in a comment, the PR body, or
     a discussion?

4. **Complexity vs need**
   - Is the abstraction level appropriate, or is there a hypothetical-future-need
     abstraction in here? (Generic helper for one caller, configuration for one
     unused option, etc.)

5. **Failure mode**
   - When this code's preconditions aren't met (bad input, missing file, network
     down), what happens? Is that the right behavior — an exception, an error
     return, a sentinel value?

Report findings to the user as a numbered list of concerns with severity
(Critical / Major / Minor) and a concrete suggestion for each. Wait for the user to
decide which to address before moving on.

## Phase 3 — Is the implementation correct?

Now read the code itself.

For each changed file, look for:

### Performance hazards

- N+1 patterns: `await` inside a loop that calls a DB / API / fs operation
- O(n²) patterns: `find` / `filter` / linear search inside a loop over the same data
- Allocations in hot paths (string concatenation in tight loops, `Array.from`/`spread`
  on large iterables)
- Repeated work that could be memoized once (regex compilation, schema parsing)

### Correctness hazards

- Off-by-one errors at boundaries (empty input, single-element input, max-size input)
- Race conditions (shared state, async without lock, observable order)
- Floating-point comparisons with `==`
- Time / locale assumptions (dates, sorting localized strings, timezone)
- Encoding assumptions (UTF-8 vs UTF-16, platform line endings)

### Type-system usage

- Are domain values typed as branded / newtype, or are they raw strings/ints
  passed around freely?
- Discriminated unions for "this OR that" instead of nullable fields for both?
- `as`/cast/`unwrap` escape hatches without a validation step before them?

### Security boundary

- All external input is validated before use (HTTP body, CLI args, file contents,
  IPC messages)
- No secrets in logs, no secrets in exception messages
- File paths are resolved with `realpath` (or equivalent) before allow-list checks —
  symlinks defeat lexical resolution
- SQL / shell interpolation uses parameterized form, not string concat

### Defensive duplication

- Is there a re-validation of something that was already validated upstream?
- Is there a `try/catch` that silently returns `null` for bugs?

### Comments and dead code

- Comments that just restate the code → delete
- Comments referencing the current PR / issue / "added for" → delete (PR body /
  git history are the right place)
- Commented-out code → delete
- `TODO` / `FIXME` left in: each one needs an owner and a tracking issue, otherwise delete

Report findings as: `[Severity] path/to/file.ext:line — description → suggested fix`.

If the diff is large, walk it in chunks (one directory, one feature area, or
~500 lines / 10 files at a time). Summarize per chunk before moving on, and let the
user decide whether to fix the chunk's findings before continuing.

## Phase 4 — Tests, docs, and changelog

A code-correct PR is not a complete PR.

### Tests

- New behavior has at least one test that would fail without the change.
- Bug fixes have a regression test that fails on the parent commit.
- Edge cases are tested separately, not just the happy path.
- Tests don't depend on each other's order or shared mutable state.
- For OSS: tests run on the public API, not internal helpers, where possible.

### Documentation

- New exports are mentioned in the README / docs, with at least one example.
- Removed or renamed exports are removed from the docs in the same PR.
- Examples in the README still compile / run after this change.

### Changelog / changeset

- User-visible changes have a changelog entry phrased from the user's perspective.
- Breaking changes are explicitly flagged.
- Pure-internal changes don't need a changeset (or mark `chore`).

### Backwards compatibility

- For pre-1.0 projects: breaking changes are allowed but should still be deliberate.
- For post-1.0 projects: breaking changes require a major bump, a deprecation
  notice on prior minor (where possible), and a migration note.

## Phase 5 — Final summary

Hand the user a short report:

```
## Review summary

Verdict: Approve / Request changes / Block

Phase 1 (motivation): OK / concerns
Phase 2 (design):     OK / N findings
Phase 3 (implementation): OK / N findings (Critical: X, Major: Y, Minor: Z)
Phase 4 (tests/docs/changelog): OK / N findings

### Must fix before merge
- ...

### Should consider
- ...

### Optional (nits)
- ...

### Maintainer follow-up after merge
- ...
```

If the verdict is **Approve**, the PR can be merged or moved out of draft.

If **Request changes**, list each item with a concrete fix and let the user address
them. After fixes, re-run Phases 3 and 4 only on the changed parts (Phase 1 and 2 are
sticky unless the design changed).

If **Block**, this PR shouldn't merge in its current form. Common reasons: rejected by
"one way to do one thing" policy, out of scope, breaking change without a path
forward, or unverifiable correctness.

## Severity definitions

| Severity   | Meaning                                                          | Action                                  |
| ---------- | ---------------------------------------------------------------- | --------------------------------------- |
| Critical   | Security hole, data corruption risk, breaking change to the public API without intent | Must fix. Cannot merge.    |
| Major      | Performance regression, design problem, missing tests for new behavior | Should fix. Skip only with explicit reason. |
| Minor      | Naming, comments, small readability issues                       | Recommend. Author's call.               |

## Mindset

- You are the **maintainer's proxy**. Defend the project's long-term cost, not the
  author's short-term convenience.
- One phase at a time. Don't batch findings into a wall of text.
- "Looks fine" is not a review. If you have nothing to say, say "no findings, moving on."
- Be specific. `[Major] src/foo.ts:42 — N+1 query → batch with bulkInsert` is useful.
  "this needs cleanup" is not.
- Don't propose new features inside a review. If you spot something adjacent that
  needs work, file an issue, don't expand the scope of the PR you're reviewing.
