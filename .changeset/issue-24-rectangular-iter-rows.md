---
'xlsx-kit': major
---

**Breaking**: `iterRows` / `iterValues` (in `xlsx-kit/worksheet`) now
iterate **rectangularly** over the populated bounding box rather than
skipping empty rows and gaps. `iterRows` yields
`(Cell | undefined)[]` (one entry per `[minCol, maxCol]` position);
`iterValues` yields `CellValue[]` with `null` filling the gaps.

Default extent switches from `MAX_ROW`/`MAX_COL` (the 1M × 16K sheet
limit) to `getMaxRow(ws)` / `getMaxCol(ws)` (the populated bounding
box). The `IterRowsOptions.valuesOnly` flag is removed — it was already
unread.

Migration:

- Aggregation callers that want populated rows only:
  `[...iterRows(ws)].filter((row) => row.some((c) => c !== undefined))`.
- Cell-by-cell streaming over populated cells only: keep using `iterCells`
  (unchanged).

Closes #24.
