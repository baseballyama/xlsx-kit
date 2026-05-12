---
'xlsx-kit': patch
---

perf: drop quadratic / linear-scan patterns on four read/write hot paths.

- The `loadWorkbook` resolver indexes each sheet's rels file once via the new `indexRelsById` helper instead of running a fresh `Array.find(r => r.id === relId)` per table / comments / drawing / chart / picture cross-reference. Worksheets with many drawings or pivot tables load in O(refs) instead of O(refs × rels).

- `containsCommentMarker` in the VML drawing classifier replaces a byte-by-byte JS loop with a single latin1 `String.indexOf` — multi-megabyte legacy VML drawings load noticeably faster.

- The streaming `iterParse` queue uses a head pointer instead of `Array#shift()`. A single SAX batch with hundreds of events (typical for a wide `<row>`) is now O(N) instead of O(N²); a stale dead-code branch in the prologue gate is gone too.

- `serializeHyperlinks` allocates rIds via a `Set` index instead of nested `Array.some()` calls. Worksheets with hundreds of hyperlinks (dashboards / link-heavy index sheets) finish save in O(N) rather than O(N²).
