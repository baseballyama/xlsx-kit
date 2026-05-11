# xlsx-kit

## 0.5.0

### Minor Changes

- [#66](https://github.com/baseballyama/xlsx-kit/pull/66) [`afecdc3`](https://github.com/baseballyama/xlsx-kit/commit/afecdc3b0b822f4d3ab3ecd16458c7a76a847f3e) Thanks [@baseballyama](https://github.com/baseballyama)! - Remove the silently-ignored `readOnly` / `keepLinks` / `keepVba` / `dataOnly` / `richText` placeholders from `LoadOptions`. They were declared on the public surface but the loader (`src/io/load.ts`) accepted them via `_opts` and dropped them on the floor, so production callers expecting `dataOnly: true` to suppress formulas — or `readOnly: true` to enable a special path — got the default behaviour instead. `LoadOptions` is now an empty type until the underlying behaviour ships; future toggles will land here once they actually do something. The `loadWorkbook(source, opts)` signature is unchanged.

### Patch Changes

- [#64](https://github.com/baseballyama/xlsx-kit/pull/64) [`b613607`](https://github.com/baseballyama/xlsx-kit/commit/b61360774e4f3b1423985ee5cf924093a991e32d) Thanks [@baseballyama](https://github.com/baseballyama)! - Treat sheet names as case-insensitive for uniqueness, matching Excel. Previously `addWorksheet(wb, 'Data')` followed by `addWorksheet(wb, 'data')` succeeded locally but produced a workbook Excel and LibreOffice refuse to open. `addWorksheet`, `addChartsheet`, `duplicateSheet`, `renameSheet`, and `pickUniqueSheetTitle` now compare titles case-insensitively. A case-only rename of the same sheet (`renameSheet(wb, 'Data', 'data')`) is allowed.

- [#51](https://github.com/baseballyama/xlsx-kit/pull/51) [`1cf8d0c`](https://github.com/baseballyama/xlsx-kit/commit/1cf8d0cc1f896366608c5d60aa40b4efc682bed9) Thanks [@baseballyama](https://github.com/baseballyama)! - Tighten the streaming I/O surface so the README's "fixed-memory" claims hold up
  in practice.

  - `toFile().toBytes().finish()` no longer re-reads the just-written file. The
    previous code called `fs.readFile(path)` from `finish()` and returned the
    full archive bytes, defeating the chunk-streamed write — a 10M-row workbook
    ended its save by reloading the entire output into memory. `finish()` now
    resolves with an empty `Uint8Array` once the underlying write stream has
    flushed; callers that need the bytes should `fs.readFile()` the path
    themselves.
  - `toFile` and `toWritable` honour write-stream backpressure: when `write()`
    returns `false`, subsequent chunks chain off a `drain` event before
    proceeding, so peak memory tracks the writable's `highWaterMark` rather
    than the producer's pace.
  - `workbookToBytes` no longer depends on `Buffer`. Browser bundles that omit
    the Node `Buffer` polyfill previously broke at `toBuffer().result()`
    because the in-memory sink ended its result with `Buffer.from(...)`. The
    helper now uses a `Uint8Array`-only sink; a regression test
    (`tests/phase-1/io/browser.test.ts`) saves a workbook with `globalThis.Buffer`
    shadowed to `undefined`.
  - The streaming read path inflates worksheet entries chunk-by-chunk. A new
    `ZipArchive.readStream(path)` returns a `ReadableStream<Uint8Array>` that
    drives fflate's `Inflate` incrementally, and `loadWorkbookStream`'s
    whole-sheet `iterRows()` feeds the SAX parser directly off that stream so
    the inflated worksheet body is never fully resident. Band queries
    (`minRow > 1`) still materialise the inflated sheet to build the row-offset
    index — that trade-off is unchanged.
  - Documentation: the `XlsxSink` / `BufferedSinkWriter` JSDoc no longer
    describes `toBytes()` as the "buffered mode" — that name was historical;
    the underlying object can either accumulate (buffered sinks) or forward
    chunks as they arrive (streaming sinks). The README also clarifies that
    the streaming reader still loads the compressed archive up front (ZIP
    needs random access to the central directory) — the win is that the
    inflated worksheet payload is never fully resident.

- [#63](https://github.com/baseballyama/xlsx-kit/pull/63) [`fa73fc5`](https://github.com/baseballyama/xlsx-kit/commit/fa73fc5a7e4c03a69eecc36acfb3a3f6482e525d) Thanks [@baseballyama](https://github.com/baseballyama)! - Tighten sheet-title validation on the streaming write path. The `createWriteOnlyWorkbook` `addWorksheet` call now applies the same rules as the buffered `addWorksheet` (no `: \ / ? * [ ]`, no leading / trailing apostrophe, not the reserved name `History`) and rejects duplicate titles case-insensitively so the streaming path can't produce a workbook Excel refuses to open.

## 0.4.0

### Minor Changes

- [`f9f273d`](https://github.com/baseballyama/xlsx-kit/commit/f9f273d390b034ac58a9d545b12ef650fa9a583a) Thanks [@baseballyama](https://github.com/baseballyama)! - Expose the full ECMA-376 axis attribute surface on `CategoryAxis` and
  `ValueAxis`. Previously the serializer emitted fixed defaults for several
  elements; these are now driven by typed fields, unblocking horizontal-bar
  reversal (`scaling.orientation: 'maxMin'`), 100 %-stacked axis caps
  (`scaling.max`), value-axis crossing rules, custom tick formatting, axis
  titles, and more.

  Newly exposed shared fields: `scaling` (`orientation`/`min`/`max`/`logBase`),
  `crosses`, `crossesAt`, `numFmt`, `majorTickMark`, `minorTickMark`,
  `tickLblPos`, `title`, `minorGridlines`. `ValueAxis` gains `crossBetween`,
  `majorUnit`, `minorUnit`. `CategoryAxis` gains `auto`, `lblAlgn`,
  `lblOffset`, `noMultiLvlLbl`. All previously-emitted defaults remain the
  output when fields are unset, so existing files are unchanged.

  New type exports: `AxisCrossBetween`, `AxisCrosses`, `AxisOrientation`,
  `AxisScaling`, `CategoryLabelAlignment`, `TickLabelPosition`, `TickMark`.

  Closes [#46](https://github.com/baseballyama/xlsx-kit/issues/46).

- [`2e5e460`](https://github.com/baseballyama/xlsx-kit/commit/2e5e4606f098ba9822bc4aaef76db324b51eeea9) Thanks [@baseballyama](https://github.com/baseballyama)! - Expose `overlap?: number` on `BarChart` (and `makeBarChart`). The serializer
  now emits `<c:overlap val="N"/>` (range -100..100) inside `<c:barChart>` when
  set, unblocking flush stacking (`overlap: 100`) and negative-space clustered
  bars. When unset, the serializer continues to emit the prior default of
  `<c:overlap val="100"/>` for `stacked` / `percentStacked` grouping so existing
  output is unchanged.

  Closes [#45](https://github.com/baseballyama/xlsx-kit/issues/45).

- [`19e8368`](https://github.com/baseballyama/xlsx-kit/commit/19e8368582c70ad89e0d6ec0265e8a7cb756ded1) Thanks [@baseballyama](https://github.com/baseballyama)! - Expose `style?: number` on `ChartSpace` (and `makeChartSpace`). The serializer
  emits `<c:style val="N"/>` (range 1..48) between `<c:roundedCorners>` and
  `<c:chart>`, selecting one of Excel's built-in "Chart Styles" gallery presets
  — the same single attribute openpyxl writes via `chart.style = N`.

  Closes [#48](https://github.com/baseballyama/xlsx-kit/issues/48).

- [`1541291`](https://github.com/baseballyama/xlsx-kit/commit/154129136198a2f22beef0a9796f2a2ba16fcaac) Thanks [@baseballyama](https://github.com/baseballyama)! - Add `DateAxis` and `SeriesAxis` types and `dateAx?` / `serAx?` slots on
  `PlotArea`. `DateAxis` carries `auto`, `lblOffset`, `baseTimeUnit`,
  `majorUnit`, `majorTimeUnit`, `minorUnit`, `minorTimeUnit` on top of the
  shared axis surface — unblocking time-series charts (`<c:dateAx>`).
  `SeriesAxis` adds `tickLblSkip` and `tickMarkSkip`, used by surface charts
  (`<c:serAx>`). The serializer emits both inside `<c:plotArea>` between the
  inferred cat/val axes and `<c:spPr>`; the parser round-trips them.

  New type exports: `DateAxis`, `SeriesAxis`, `TimeUnit`.

- [`0708aa8`](https://github.com/baseballyama/xlsx-kit/commit/0708aa81106db5f4276d1d214fc5714e25996fb3) Thanks [@baseballyama](https://github.com/baseballyama)! - Add `Layout` / `ManualLayout` types and expose `layout?: Layout` on
  `ChartTitle`, `PlotArea`, and `Legend`. The serializer emits
  `<c:layout><c:manualLayout>` with `layoutTarget`, `xMode` / `yMode` /
  `wMode` / `hMode`, and `x` / `y` / `w` / `h` when set, falling back to the
  existing empty `<c:layout/>` placeholder when unset — so output is unchanged
  for charts that don't configure manual layout. Parser round-trips both
  forms.

  New type exports: `Layout`, `LayoutMode`, `LayoutTarget`, `ManualLayout`.

- [`0989eec`](https://github.com/baseballyama/xlsx-kit/commit/0989eec45f879d05a7707da8402fd734f4a3208b) Thanks [@baseballyama](https://github.com/baseballyama)! - Expose per-point `dPt?: DataPoint[]` on `BarSeries` (used by bar / line /
  area / pie / doughnut / radar / stock / surface), `ScatterSeries`, and
  `BubbleSeries`, with the new `DataPoint` type carrying `idx`,
  `invertIfNegative?`, `marker?`, `bubble3D?`, `explosion?`, and `spPr?`.
  The serializer emits `<c:dPt>` children between the series'
  `<c:marker>`/`<c:spPr>` and `<c:dLbls>` per ECMA-376 sequence — unblocking
  per-slice colours on pie / doughnut charts, per-bar colours on single-series
  bar charts, and per-point styling on line / scatter / bubble.

  Closes [#44](https://github.com/baseballyama/xlsx-kit/issues/44).

- [`7f9e143`](https://github.com/baseballyama/xlsx-kit/commit/7f9e1430c32ad685b14e382c1a17abac41f24b4f) Thanks [@baseballyama](https://github.com/baseballyama)! - Add `invertIfNegative?: boolean` and `explosion?: number` to `BarSeries`
  (used by bar / line / area / pie / doughnut / radar / stock / surface) and
  `invertIfNegative?: boolean` to `BubbleSeries`. The serializer emits
  `<c:invertIfNegative>` and `<c:explosion>` between `<c:spPr>` and `<c:dPt>`
  per ECMA-376 sequence — unblocking per-series colour inversion on negative
  values and pie/doughnut slice explosion at the series level (in addition to
  the per-point `DataPoint.explosion`).

- [`ffa777c`](https://github.com/baseballyama/xlsx-kit/commit/ffa777caaddbf00f2cffc7fedce8d021a2a584f6) Thanks [@baseballyama](https://github.com/baseballyama)! - Expose `marker?: Marker` on `LineSeries` and `ScatterSeries` (with the new
  `Marker` / `MarkerSymbol` types). The serializer emits `<c:marker>` between
  the series' `<c:spPr>` and `<c:dLbls>` per ECMA-376 sequence, carrying
  `<c:symbol>`, `<c:size>`, and an optional nested `<c:spPr>` for marker
  fill / line colour — matching openpyxl's `series.marker = Marker(...)`.

  Closes [#47](https://github.com/baseballyama/xlsx-kit/issues/47).

- [`70a2f17`](https://github.com/baseballyama/xlsx-kit/commit/70a2f17bfe6a1106df04702de5e11e2ac16cd596) Thanks [@baseballyama](https://github.com/baseballyama)! - Extend `StockChart.hiLowLines` and `StockChart.upDownBars` to accept a
  detailed object form in addition to the existing boolean flag. The
  detailed form lets callers style the lines (`HiLowLines.spPr`) and the
  up/down bars (`UpDownBars.gapWidth` + `upBars.spPr` + `downBars.spPr`)
  with per-element shape properties.

  The boolean form (`hiLowLines: true`) keeps its existing meaning and
  output, so existing callers are unaffected. Parser round-trips both
  forms, picking the boolean form when no detail is found.

  New type exports: `BarFrame`, `HiLowLines`, `UpDownBars`.

- [`7cd181c`](https://github.com/baseballyama/xlsx-kit/commit/7cd181c9276dfe8c37675d3c4b1c77e020b82b64) Thanks [@baseballyama](https://github.com/baseballyama)! - Add `view3D?: View3D` and `floor?` / `sideWall?` / `backWall?` (typed
  `SurfaceFrame`) to `ChartSpace` (and `makeChartSpace`). The serializer
  emits `<c:view3D>` (with `rotX`, `rotY`, `depthPercent`, `hPercent`,
  `rAngAx`, `perspective`) and `<c:floor>` / `<c:sideWall>` / `<c:backWall>`
  (with `thickness` and `spPr`) between `<c:autoTitleDeleted>` and
  `<c:plotArea>` per ECMA-376 sequence — unblocking real 3-D chart viewpoints
  and wall styling for `bar3DChart` / `line3DChart` / `pie3DChart` /
  `area3DChart` / `surface3DChart`.

## 0.3.1

### Patch Changes

- [#41](https://github.com/baseballyama/xlsx-kit/pull/41) [`a04f645`](https://github.com/baseballyama/xlsx-kit/commit/a04f6459113810c638adf4247d1a201e34123d1c) Thanks [@baseballyama](https://github.com/baseballyama)! - Relax `engines.node` from `>=24.15.0` back to `>=22.0.0` so the published
  package installs on every active Node LTS line (22.x, 24.x) plus current
  (26.x), matching the CI matrix. 0.3.0 inadvertently shipped a Node 24+
  floor that excluded the still-supported 22.x LTS; this restores broader
  LTS coverage. The library does not rely on any Node 24-only API.

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
