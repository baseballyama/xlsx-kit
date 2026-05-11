---
name: review-response
description: Use when responding to GitHub PR review comments. Covers fetching unresolved threads only, deciding "fix" vs "decline" per comment with the user, replying with the commit hash, and feeding back lessons to the project's skill files. Invoked when the user says "respond to review", "address review comments", or pastes a PR with review feedback.
---

# Responding to review comments

GitHub review comments are a back-and-forth, not a checklist. The right loop:
**fetch unresolved → decide per comment with the user → fix or decline → reply → repeat**.

## Principles

| Principle                  | Detail                                                                |
| -------------------------- | --------------------------------------------------------------------- |
| Only unresolved threads    | If the reviewer marked it resolved, don't reopen it.                  |
| One at a time              | Even with many comments, address them sequentially.                   |
| Always reply               | Every comment gets a response — fix or explain.                       |
| Reference the commit hash  | Replies that say "fixed" without a hash leave the reviewer guessing.  |
| Decline with a reason      | "Won't fix" is fine; silent or hand-wavy "won't fix" is not.          |
| Get user approval first    | Decide the plan with the user before posting anything visible.        |

## Step 0 — Fetch unresolved threads

GitHub's REST endpoint returns *all* review comments — including resolved ones. Use
GraphQL to filter for unresolved.

```bash
# All unresolved review threads on a PR
gh api graphql -f query='
  query($owner: String!, $repo: String!, $pr: Int!) {
    repository(owner: $owner, name: $repo) {
      pullRequest(number: $pr) {
        reviewThreads(first: 100) {
          nodes {
            isResolved
            isOutdated
            comments(first: 20) {
              nodes {
                id
                databaseId
                author { login }
                body
                path
                line
                url
              }
            }
          }
        }
      }
    }
  }
' -f owner=<owner> -f repo=<repo> -F pr=<pr>
```

Filter the result client-side for `isResolved == false`. Each thread's first comment
is the original; the rest are replies. You're addressing the original, but read the
whole thread — the reviewer may have already self-answered, or another contributor
may have responded.

Also fetch:

- **General PR comments** (`gh api repos/{owner}/{repo}/issues/{pr}/comments`) — these
  are not "review" comments but are part of the conversation.
- **Review summary bodies** (`gh api repos/{owner}/{repo}/pulls/{pr}/reviews`) — when
  a reviewer leaves a body on the overall review, not just inline.

## Step 1 — Plan with the user

Before writing any code or replies, present the plan as a table:

| #   | Where                  | Comment (summary)               | Decision               | Plan                                        |
| --- | ---------------------- | ------------------------------- | ---------------------- | ------------------------------------------- |
| 1   | src/foo.ts:42          | Possible N+1 in this loop       | **Fix**                | Replace `find` with `Map`-based lookup.     |
| 2   | src/bar.ts:17          | Should add a comment here       | **Decline (won't fix)** | The function name already states this. Reply explaining. |
| 3   | docs/README.md:108     | Typo: "lenght" → "length"       | **Fix**                | One-character change.                       |

Decision options:

- **Fix** — the comment is correct; we'll change the code.
- **Defer** — fair point but out of scope for this PR; we'll file a follow-up issue.
- **Decline** — disagreement (technical mistake in the comment, already addressed,
  out of project scope, etc.). Must include the *reason* you'd give in the reply.

**Wait for the user's approval** of this table before doing anything visible. If the
user revises decisions, follow the revision.

## Step 2 — Address comments one at a time

For each row in the agreed table:

1. **Pull latest first** — `git pull --rebase` so you don't fix on top of stale
   state.
2. **Make the change** (only if "Fix"). Keep the diff scoped to that one comment.
3. **Run the minimum gates** — format, lint, type check. Don't run the full test
   matrix per comment; do that once at the end. (See `run-check-and-test`.)
4. **Commit** with a message that references the comment:

   ```
   fix: avoid N+1 in foo loop (review #PR123-c456)
   ```

   The hash is what the reply will reference, so this is the load-bearing artifact —
   make the message clear.

5. **Push** to the PR branch.
6. **Reply on GitHub** with the appropriate template.

## Reply templates

**When you fixed it:**

```markdown
Thanks for catching this!

[One sentence: what you changed and why it addresses the comment.]

Fixed in <commit-hash>.
```

**When you're declining:**

```markdown
Thanks for the suggestion!

[One paragraph: why this isn't going to change. Be concrete: cite the existing
behavior, the project policy, or the technical reason.]

- [Specific point 1]
- [Specific point 2]

Going to leave the implementation as is for [reason]. Happy to revisit if I'm
missing context.
```

**When you're deferring:**

```markdown
Good point — this is real but out of scope for this PR. Filed as #<issue-number>
to keep the conversation tracked.
```

## Reply API

```bash
# Reply on a line-level review comment (creates a threaded reply)
gh api repos/<owner>/<repo>/pulls/<pr>/comments/<comment-id>/replies \
  -f body="<message>"

# Reply on a general PR (issue-style) comment
gh api repos/<owner>/<repo>/issues/<pr>/comments \
  -f body="<message>"
```

Use the `databaseId` from the GraphQL fetch as `<comment-id>` for line-level replies.

## Step 3 — Final checks once all comments are addressed

After the last comment:

- Run the full quality gate (`run-check-and-test`).
- Re-read your own diff. The cumulative result of many small fixes can have
  unintended interactions.
- Push once more if anything came up. Make a note in the PR thread that everything
  is addressed and the PR is ready for re-review.

## Feedback loop: AI-bot review comments

When the comment author is an AI bot (CodeRabbit, Codex, similar), there's an extra
step worth doing.

### When the bot was right and you fixed it

The same class of issue may exist elsewhere or come up again. Consider:

1. Search the project's skill files (`.claude/skills/`) for the pattern.
2. If it's not already documented, add it to the most relevant skill (e.g.,
   coding-principles, backend, frontend). Keep it short — one example is enough.
3. Mention to the user that you updated the skill, so they can confirm.

### When the bot was wrong and you declined

The same false positive will probably come back on the next PR. Consider:

1. Add a note to the relevant skill explaining *why* this pattern is fine in this
   project. The skills act as the project's "knowledge base" for AI tools, so
   recording the rationale once reduces repeat false positives.
2. Mention to the user that you added context.

In both cases, **don't update skill files without telling the user first.** This is a
project-level change, not a per-PR change.

## What not to do

- Don't reply "fixed" without a commit hash.
- Don't push a single squash commit at the end of the conversation that flattens all
  fixes — reviewers can no longer tell which commit addressed which comment.
- Don't `--force-push` during active review (rebasing on top of `main` mid-review
  invalidates threads). If you must rebase, tell the reviewer first.
- Don't argue. If you disagree, decline with a reason. If they push back with a real
  argument, change your mind and fix it. If they push back with the same vibe, say
  "let's get [maintainer] to weigh in."
- Don't @-mention people who weren't already in the thread.
