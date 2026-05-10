# xlsx-kit

## 0.3.0

### Minor Changes

- [#32](https://github.com/baseballyama/xlsx-kit/pull/32) [`87a0051`](https://github.com/baseballyama/xlsx-kit/commit/87a005104f4d54b5fd0a1a747acc515d6cf9171e) Thanks [@baseballyama](https://github.com/baseballyama)! - **Breaking**: `iterRows` / `iterValues` (in `xlsx-kit/worksheet`) now
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

  Closes [#24](https://github.com/baseballyama/xlsx-kit/issues/24).

- [#31](https://github.com/baseballyama/xlsx-kit/pull/31) [`9297c46`](https://github.com/baseballyama/xlsx-kit/commit/9297c46aed63f566b081db97fc1f84ca9a24b3c0) Thanks [@baseballyama](https://github.com/baseballyama)! - Extend `cellValueAsString` (in `xlsx-kit/cell`) with optional
  `dateFormat` / `emptyText` overrides and add a sibling
  `cellValueAsPrimitive` that maps a `CellValue` to the most natural JS
  primitive (`string | number | boolean | Date | null`) without forcing a
  single target type. Closes [#25](https://github.com/baseballyama/xlsx-kit/issues/25).

- [#30](https://github.com/baseballyama/xlsx-kit/pull/30) [`78d04fd`](https://github.com/baseballyama/xlsx-kit/commit/78d04fd8cee4f9b2b0991e14f471ede39adbef2f) Thanks [@baseballyama](https://github.com/baseballyama)! - Add `workbookToBuffer` to `xlsx-kit/node`. One-shot Node-flavored helper
  that returns a `Buffer` directly, paralleling the existing `fromBuffer`
  source. Closes [#28](https://github.com/baseballyama/xlsx-kit/issues/28).

## 0.2.0

### Minor Changes

- [`b36ca45`](https://github.com/baseballyama/xlsx-kit/commit/b36ca453b08c91981baac42b3b5bc4aeeeef6ec0) Thanks [@baseballyama](https://github.com/baseballyama)! - Hardening and docs release.

  - Add a 3-tier ECMA-376 conformance validator and broaden conformance coverage to the writer surface, real-world fixtures, and fast-check property tests.
  - Add `knip` to CI to keep the public export surface tight; prune unused exports flagged by it.
  - Refresh the docs site: redesigned landing and docs UI with a new typography system, new logo and favicons, and new "Why xlsx-kit" / comparison / motivation sections in the README.
  - Tighten release / dependency automation: pin dependencies, drop EOL Node 18/20 from the test matrix and add Node 26, bump the project Node engine to 22.22.2.
