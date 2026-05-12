---
'xlsx-kit': patch
---

refactor: collapse `describeWorkbook` into a single pass, harden `getRowValues`, and document `normalizePath`'s `..` handling.

`describeWorkbook` used to walk every cell three times — once for `getWorkbookStats`, once for `getWorkbookCellsByKind`, and once again for the per-sheet counts. The new implementation fuses all three into one pass and shares the value-kind classifier with `countCellsByKind` via the newly exported `classifyCellValue` (so the two never drift on edge cases like `Date` vs `duration`).

`getRowValues` no longer derives `maxCol` via `Math.max(...rowMap.keys())`. The spread is limited by V8's argument-count cap (~125 k), which `getRowValues` would silently blow past on a dense row near MAX_COL — the call now uses a linear scan to derive the max.

`normalizePath` picks up an explanatory comment for the `..`-when-`out`-is-empty case, so the next reader doesn't have to second-guess whether path-traversal is possible (it isn't — the archive lookup catches escape attempts naturally).
