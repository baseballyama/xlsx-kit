---
name: pr-workflow
description: Use when opening a pull request. Covers branch naming, commit hygiene, PR title and body conventions, and what must be true before a PR leaves draft. Invoked when the user says "open a PR", "create a PR", "submit", or asks to push and PR a branch.
---

# PR workflow

The shape of a PR communicates the change to maintainers and to a future archaeologist
reading `git log`. Get it right at write time — the cost of fixing a confusing PR
later is much higher than getting it right once.

## Before you open the PR

1. **Branch is up to date with the base.** Rebase onto `main` (or whatever the base
   is) so the PR diff is clean. Don't merge `main` into your feature branch as a
   habit; rebase.
2. **All quality gates pass locally.** Use the `run-check-and-test` skill — typecheck,
   lint, format, test, build. CI will catch you if you skip this; just do it locally.
3. **Each commit is meaningful.** Squash WIP commits. The git log should read as a
   sequence of "I did this, then I did this, then I did this," not "fix typo / fix
   again / really fix this time."
4. **The change has an issue or design discussion to point at**, unless it's
   trivially small (typo, comment, single-line bugfix). For non-trivial features, the
   discussion happens *before* the PR, not in PR review.

## Branch name

```
<type>/<short-slug>
```

Conventional types match the commit prefixes below: `feat`, `fix`, `refactor`, `docs`,
`test`, `chore`, `perf`, `ci`, `build`. If the work resolves a specific issue, include
the number: `fix/issue-123-empty-sheet-crash`.

## PR title

Use [Conventional Commits](https://www.conventionalcommits.org/) prefixes — they
drive the changelog and make scanning history easy.

```
<type>(<optional-scope>)<!?>: <imperative summary>
```

| Prefix     | When to use                                                  |
| ---------- | ------------------------------------------------------------ |
| `feat`     | New user-visible capability                                  |
| `fix`      | Bug fix                                                      |
| `refactor` | Internal change with no user-visible behavior change         |
| `perf`     | Performance improvement                                      |
| `docs`     | Documentation only                                           |
| `test`     | Tests only                                                   |
| `build`    | Build system, dependencies                                   |
| `ci`       | CI configuration                                             |
| `chore`    | Tooling, repo plumbing                                       |
| `revert`   | Reverting a previous commit                                  |

Add `!` for breaking changes (e.g., `feat(api)!: rename loadWorkbook to readWorkbook`),
or note `BREAKING CHANGE:` in the body.

**Title is from the user's perspective**, not yours.

```
✅ feat: add streaming loader for files larger than memory
✅ fix: empty sheet names crashed loadWorkbook (#123)
✅ perf(parser): avoid quadratic scan when resolving named ranges

❌ refactor SheetReader to use new helper       (no prefix, what does the user gain?)
❌ feat: implement getCellValue helper           (does the user know what that solves?)
❌ fix: change line 42 in worksheet.ts           (code-level, not user-level)
```

## PR body

The body must include the anchors used by the `template-compliance` workflow.
**Do not remove the HTML comment markers.**

```markdown
<!-- pr-template:v1 -->

## Summary

<!-- One short paragraph: what this PR changes and why a user / maintainer should care.
     Not a commit-by-commit walkthrough. -->

## Motivation

<!-- The problem this solves. Link to the issue or design discussion that proposed it.
     If there is none, explain why this was opened without one. -->

Closes #<issue-number>  <!-- if applicable -->

## Changes

<!-- Bullet list of user-visible changes. New exports, removed exports, behavior
     changes, schema changes. Skip purely internal refactors. -->

- ...

## Testing

<!-- How you verified this works. New tests, existing tests, manual repro steps.
     Include a command a reviewer can run to reproduce. -->

- ...

## Breaking changes

<!-- "None" if not breaking. Otherwise: what breaks, who is affected, what they need
     to do, deprecation plan. -->

None

## Checklist

- [ ] I have read CLAUDE.md and followed the project's conventions.
- [ ] I have added or updated tests for the change.
- [ ] I have added or updated documentation where user-visible behavior changed.
- [ ] If this is a breaking change, I have added a changeset / CHANGELOG entry and
      flagged it above.
- [ ] I have re-read my own diff and removed dead code, debug prints, and stale comments.
- [ ] If I used an LLM to draft this PR, I have verified each change myself, the PR
      represents real work that warrants a maintainer's review, and I am willing to
      defend each line in review.

<!-- pr-template:end -->
```

## Commit messages

Same prefix system as PR titles. The commit body, when present, explains **why** the
change was needed and **what tradeoffs** were considered. Save "what" for the diff —
it's already there.

```
fix: handle empty sheet names in loadWorkbook (#123)

Empty <sheet name=""> attributes in the worksheet relationships file
caused the parser to misalign sheet indices, throwing a generic
TypeError. ECMA-376 §18.2.20 explicitly allows empty names, so we now
treat them as a valid (but unindexed) sheet rather than rejecting.

Closes #123
```

## Draft vs ready

- **Draft** while the change is in flux — CI is allowed to be red, the description is
  rough, you're still iterating.
- Move out of draft only when:
  - All checks pass.
  - The PR body is final (the maintainer will read it next).
  - You've reviewed your own diff line by line.

If you ping a reviewer on a draft, mention what feedback you specifically want — "is
the API shape ok?" reads very differently from "please review."

## Self-review before requesting reviewers

Read your own diff as if you were the reviewer:

- Are there commented-out blocks? Print statements left in? Generated/scratch files?
- Is each new identifier necessary, or did you keep an intermediate variable for no reason?
- Does every comment still match the code next to it after your edits?
- For a stranger picking this up cold: is the *why* visible in the PR body, the
  commit messages, or the comments — at least one of those?

If anything is off, fix it before tagging reviewers. Reviewers' time is the most
expensive thing in this loop; a self-review pass costs you minutes and saves them
hours.

## After opening

- Watch CI. If something goes red, fix it before reviewers spend time on the diff.
- Respond to review comments using the `review-response` skill — one comment at a
  time, with the commit hash that addresses it.
- **Don't force-push during active review** unless you're squashing the final commits
  on request. Force-pushes invalidate review threads and make it hard for the
  reviewer to see incremental changes.

## When the change is internal-only

If this PR doesn't change anything user-visible (pure refactor, internal tooling,
test cleanup):

- Use the `refactor` / `chore` / `test` prefix.
- Don't add a changeset (or mark it `chore` with no changelog entry).
- The "Breaking changes" section is "None" — keep it.

If you're unsure whether the change is user-visible, ask: "would a user upgrading
across this commit notice anything?" If yes, it's user-visible.
