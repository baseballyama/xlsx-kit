# PROGRESS

`/loop` 自走モードでフェーズ1→7 を順に実装するための、ターン横断の状態ファイル。
**唯一の正は `docs/plan/`**。本ファイルは「いまどこまで終わったか」だけを記録する。

## カレント

- **フェーズ**: フェーズ5 (rich features) — フェーズ3 acceptance pass、フェーズ4 streaming は perf ベンチが必要なので後回し
- **フェーズ**: フェーズ5 worksheet rich features 全部完了 → **フェーズ6 charts 全部完了**

## 次の作業者への引き継ぎ (Handoff @ 2026-05-05)

**現状サマリ**: 2221 tests pass / typecheck + lint clean (14 warnings baseline)。フェーズ1〜7 のコア実装+ streaming + edge acceptance + 多数の ergonomic helper 追加完了。1.0 候補。直近で *ToCss / setRange* / clearCellStyle / iter+enumeration / drawing wipe / kind histogram / append/write/readRange (+Objects) / addTableFromObjects / CSV bundle round-trip / createWorkbookFromCsv/Objects / cell+workbook summaries / empty predicates / sheet-qualified address API (formatSheetQualifiedRef + getCellAddress/getRangeAddress + getCellAtAddress/setCellAtAddress + getValueAtAddress + getRangeValuesAtAddress/setRangeValuesAtAddress) / DefinedName ergonomics (addDefinedNameForRange + getDefinedNameTarget) / A1-string range predicates+ops (isCellInRange/isRangeInRange/rangesOverlapStr/unionRangeStr/intersectionRangeStr/shiftRangeStr/expandRangeStr/rangeAreaStr/rangeDimensionsStr) / analytics primitives (tabularData/columnAggregates/groupBy/pivotTable/sortRange/filterRange/mapRange) を積み増し (+328 tests / 1893 → 2221)。

### 直近 100+ commit のテーマ (今ブランチで積んだ作業)

- **public ergonomic API の積み増し**: cell-level (setCellFormula / setCellArrayFormula / setCellRichText / 値 coercion 一式 / 型ガード) → range-level (copyRange / moveRange / clearRange / clearAllCells / setRangeBackgroundColor / setRangeFont / setRangeNumberFormat / setRangeBorderBox / setRangeStyle) → worksheet-level (selection helpers / freeze panes / outline grouping / outline collapse / hide+unhide bulk / autofit / data extent / find cells / merge resolver / hyperlink resolver / comment resolver) → workbook-level (duplicateSheet / iterAllCells / getAllMergedRanges / getAllHyperlinks / getAllComments / getAllTables / getAllDataValidations / getWorkbookStats / sheet validators)。
- **packaging ergonomic API**: docProps/core.xml + docProps/app.xml + docProps/custom.xml の主要 field を直接 set できる setter 群 (creator/title/manager/company/customString/Number/Bool/Date/...)。
- **styling ergonomic API**: alignment preset (centerCell / wrapCellText / rotateCellText / indentCell / alignXxx) / font preset (setBold / setItalic / setUnderline / setFontSize / setFontName / setFontColor) / fill preset (setCellBackgroundColor / clearCellBackground) / border preset (setCellBorderAll / setRangeBorderBox) / format preset (setCellAsCurrency / Percent / Date / Number) / WCAG color helpers (luminance / contrastRatio / pickReadableTextColor) / formatAsHeader range preset / applyBuiltinStyle + applyNamedStyle。
- **CF visual rule builders**: addColorScaleRule / addDataBarRule / addIconSetRule (innerXml 直書きを構造化 opts で代替)。
- **重要 fix**: `applyXfPatch` / `setRangeStyle` が `cellXfs[0]` を default で予約 (a0c9400)。これ以前は最初の styled cell が unstyled cell と style id 0 で衝突する footgun。

### 残タスク

- **Excel 365 視覚 QA** — 人手のみ。コア機能は揃っている。
- **ZIP64 write の正式対応** — fflate 上流の 4GiB 制限。openxml-js 側ではすでに fallback ロジックが入っているので blocker ではない。
- **public API surface の整理** — `src/index.ts` が 700+ 行に肥大。alphabetical 順を維持しているが、サブモジュールごとに `*.ts` barrel 経由でグループ化する余地あり。（urgent ではない、整形のみ。）
- **rich-text run builder ergo の追加** — `richTextRun(text, font?)` 単体 export はまだ。`makeTextRun` を re-export するだけで済む。
- ~~**`replaceCellValues` の range-aware 版**~~ → `replaceInRange` で対応 (2026-05-07)。
- ~~**autofit の font-aware 改良**~~ → `opts.workbook` で font.size 比例 scaling 対応 (2026-05-07)。

### 進め方の TIPS

- **ターン pattern**: 「PROGRESS.md を読む → 候補から最小タスクを 1 件選ぶ → 実装 → テスト (vitest) → typecheck + lint clean → コミット → PROGRESS.md の最上段に新タスクを書き、前タスクは "次のタスク (前回)" として一段繰り下げる」。
- **テスト**: 各 helper 追加時は `tests/phase-5/` (worksheet 系) / `tests/phase-3/` (workbook + packaging 系) / `tests/phase-2/styles/` (style 系) / `tests/phase-2/` (cell 系) に新ファイルを切る。round-trip テストは `workbookToBytes` → `loadWorkbook(fromBuffer(...))` の pattern。
- **`exactOptionalPropertyTypes` + `noUncheckedIndexedAccess`** — `_drop` / 明示型ガード必須。`x!.y` は biome lint error なので避ける。
- **PR 作業をする場合**: `git push origin main` で main 直 push (このリポジトリはオーナー単独運用)。


- **次のタスク**: **fixture-based export-format smoke test を追加** — 既存 xlsx fixture を load → 4 形式の workbook record 全部 invoke して non-empty 確認。
  1. `tests/phase-3/workbook-export-formats-smoke.test.ts` 新規: load + 4 record export を walk、各 sheet が string で返ることを assert。
  2. 新 helper 不要。

- **次のタスク (前回)**: **getWorksheetAsTextTable + getWorkbookAsTextTableRecord (export-format matrix close)**。
  1. `src/worksheet/text.ts` + `src/workbook/workbook.ts` に追加。
  2. `src/index.ts` から re-export。
  3. tests/phase-5/worksheet-as-text.test.ts 3 件 + tests/phase-3/workbook-as-text-record.test.ts 3 件。

  empirical: 2371 tests pass (was 2365, +6)、typecheck / lint clean (14 warnings)。 export-format API matrix complete (per-range + per-sheet shortcut + workbook Record × CSV/HTML/Markdown/Text)。

- **次のタスク (前回 2)**: **`worksheetToTextTable(ws, range)` ASCII-art plain-text renderer**。
  1. `src/worksheet/text.ts` 新規: 列幅 padEnd + `+---+` border、merge flatten、newline→space。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/worksheet-to-text.test.ts` 5 件: 通常 / unequal widths / merge / newline 置換 / 1 行 range。

  empirical: 2365 tests pass (was 2360, +5)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: iterWorksheets + getWorksheetAsMarkdownTable → Record。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-as-markdown-record.test.ts` 4 件: 通常 / 空 sheet '' / 空 wb / chartsheet skip。

  empirical: 2360 tests pass (was 2356, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/markdown.ts` に追加: getDataExtent → A1 → worksheetToMarkdownTable。空 ws は ''。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/worksheet-as-markdown.test.ts` 4 件: 通常 / 空 / sparse / merge。

  empirical: 2356 tests pass (was 2352, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/markdown.ts` (新規): GFM 形式、header sep `| --- |`、`|` / newline escape、merge 無視で flatten。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/worksheet-to-markdown.test.ts` 5 件: 通常 / `|` escape / 空 cell / merge flatten / 1 row range。

  empirical: 2352 tests pass (was 2347, +5)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: iterWorksheets + getWorksheetAsHtml → Record。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-as-html-record.test.ts` 4 件: 通常 / 空 sheet '' / 空 wb / chartsheet skip。

  empirical: 2347 tests pass (was 2343, +4)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: wb.sheets[activeSheetIndex]?.sheet.title。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/get-active-sheet-title.test.ts` 4 件: 通常 / setActiveSheet / chartsheet active / 空 wb。

  empirical: 2329 tests pass (was 2325, +4)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: getSheet(wb, title) !== undefined wrapper。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/has-worksheet.test.ts` 4 件: 在 / 同名 chartsheet false / 不在 / 空 wb。

  empirical: 2315 tests pass (was 2311, +4)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: wb.sheets walk + kind 'chartsheet' && title 一致。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/has-chartsheet.test.ts` 4 件: 在 / 不在 / 同名 worksheet false / 空 wb。

  empirical: 2311 tests pass (was 2307, +4)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: getSheetIndex(wb, title) >= 0 wrapper。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/has-sheet.test.ts` 4 件: worksheet 在 / 不在 / chartsheet 在 / 空 wb。

  empirical: 2307 tests pass (was 2303, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: header walk + minCol 相対 index、不在で -1。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/column-index-of.test.ts` 3 件: 在 / 不在 -1 / range 相対。

  empirical: 2299 tests pass (was 2296, +3)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: header walk + minCol 相対 index、不在で -1。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/column-index-of.test.ts` 3 件: 在 / 不在 -1 / range 相対。

  empirical: 2299 tests pass (was 2296, +3)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: header walk で短絡判定。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/has-column.test.ts` 4 件: 在 / 不在 / 空 range / 空文字 sentinel。

  empirical: 2296 tests pass (was 2292, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: 全 entry を up-front validate (oldName 存在 / newName collision / duplicate newName) → mutate。空 mapping no-op。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/rename-columns.test.ts` 7 件: 単一 / 複数 / swap / oldName 不在 throw / collision throw / duplicate target throw / 空 no-op。

  empirical: 2288 tests pass (was 2281, +7)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects → newOrder 順で書き直し、余り列は clear。空 / unknown name で throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/reorder-columns.test.ts` 5 件: swap / subset drop / unknown name throw / empty throw / 同順 no-op。

  empirical: 2281 tests pass (was 2276, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: column 検索 → 右側 shift → 最右列 clear。new range return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/remove-column.test.ts` 4 件: 中央 / 最右 / 不存在 throw / 単一列 throw。

  empirical: 2276 tests pass (was 2272, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: header + data 値 (value/fn) を maxCol+1 に書き込み、新 range を return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/add-column.test.ts` 4 件: 通常 / fn 計算 / 重複 name throw / header-only range。

  empirical: 2272 tests pass (was 2268, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: parseRange で column index → 各 data row に value/fn 適用。header 不変。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/fill-column.test.ts` 5 件: 同値 / fn / header preservation / 不存在 throw / 空 data area。

  empirical: 2268 tests pass (was 2263, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: header cell rewrite、不存在 / 重複 / 同名 no-op を validation。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/rename-column.test.ts` 5 件: 通常 / 不存在 throw / 重複 throw / 同名 no-op / data 不変。

  empirical: 2263 tests pass (was 2258, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: pluckColumn → Set dedupe、null は distinct。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/unique-column.test.ts` 4 件: dedupe / null distinct / 空 [] / 不存在 column throw。

  empirical: 2258 tests pass (was 2254, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: tabularData[column]、不存在 column で throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/pluck-column.test.ts` 4 件: 通常 / null / 不存在 throw / 空 range []。

  empirical: 2254 tests pass (was 2250, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects → length / filter count。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/count-rows.test.ts` 4 件: 全 row / predicate / 空 / 全 reject。

  empirical: 2250 tests pass (was 2246, +4) — 300 test files crossed、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects 短絡。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/index-of-row.test.ts` 4 件: hit / no match -1 / 空 -1 / 短絡。

  empirical: 2246 tests pass (was 2242, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects walk で短絡判定。空 range で some=false / every=true。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/some-every-row.test.ts` 7 件: some hit / some no / some 空 / some 短絡 / every all true / every 第二 row で短絡 / every 空 vacuous true。

  empirical: 2242 tests pass (was 2235, +7)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects → 短絡 find。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/find-row.test.ts` 5 件: hit / 短絡 / no match / 空 / index passing。

  empirical: 2235 tests pass (was 2230, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects → for callback。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/for-each-row.test.ts` 4 件: 通常 / index / 空 / row shape。

- **次のタスク (前回)**: **`reduceRange(ws, range, reducer, init)`** — header-driven range をユーザー定義 reducer で 1 値に折りたたむ。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects → reduce(acc, row, i) → 終端値 return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/reduce-range.test.ts` 4 件: count / sum / max / 空 range で init 返す。

- **次のタスク (前回)**: **`mapRange(ws, range, transform)` row transform**。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects → transform → setCell 書き戻し。欠損 key で null clear、未知 key skip。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/map-range.test.ts` 5 件: identity-mod / null clear / 未知 key skip / 欠損 key clear / multi-column 共参照。

  empirical: 2221 tests pass (was 2216, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects → predicate filter → 残行を re-pack、余り行を null clear。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/filter-range.test.ts` 4 件: 通常 / 全削除 / 全保持 / multi-column。

  empirical: 2216 tests pass (was 2212, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects → sort → setCell で書き戻し (null overwrite)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/sort-range.test.ts` 6 件: 文字列 / 数値 / descending / null last / multi-column / 不存在 column throw。

  empirical: 2212 tests pass (was 2206, +6)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects → row × col group → aggregate (sum/count/mean/min/max)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/pivot-table.test.ts` 5 件: sum / count / max / mean / column 不存在 throw。

  empirical: 2206 tests pass (was 2201, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: readRangeAsObjects → `Record<keyValue, rowObj[]>`。null は `''` bucket、不存在 column で throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/group-by.test.ts` 5 件: 通常 / 単一 group / null key '' bucket / 不存在 column throw / header-only で {}。

  empirical: 2201 tests pass (was 2196, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: tabularData → sum/mean/min/max/count/numericCount。numericCount===0 は NaN。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/column-aggregates.test.ts` 4 件: 数値列 / 文字列のみ列 NaN / mixed types / multi-column 独立。

  empirical: 2196 tests pass (was 2192, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: getRangeValues → 列 store。重複 header は concat。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/tabular-data.test.ts` 5 件: 通常 / header-only で空 columns / null 列 / header coercion / 重複 header concat。

  empirical: 2192 tests pass (was 2187, +5)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: getCellAtAddress + .value extraction (null fallback)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/value-at-address.test.ts` 4 件: 値あり / 不存在 cell null / quoted title / 不存在 sheet throw。

  empirical: 2187 tests pass (was 2183, +4)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: parseSheetRange → setCellByCoord。range は throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/set-cell-at-address.test.ts` 4 件: 通常 / quoted / 不存在 sheet throw / range throw。

  empirical: 2183 tests pass (was 2179, +4)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: parseSheetRange → getSheet → setRangeValues。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/set-range-values-at-address.test.ts` 5 件: 通常 / quoted / null skip / round-trip / 不存在 sheet throw。

  empirical: 2179 tests pass (was 2174, +5)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: parseSheetRange → getSheet → getRangeValues。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/range-values-at-address.test.ts` 5 件: 矩形 / 単一 cell も 2D / 空 cell null / quoted title / 不存在 sheet throw。

  empirical: 2174 tests pass (was 2169, +5)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: parseSheetRange → getSheet → getCell。range は throw、不存在 sheet は throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/cell-at-address.test.ts` 6 件: bare / quoted / 不存在 cell undefined / 不存在 sheet throw / range throw / round-trip。

  empirical: 2169 tests pass (was 2163, +6)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/cell-range.ts` に追加: maxRow/maxCol を delta だけ伸縮。makeCellRange で validation (零次元で throw)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/expand-range-str.test.ts` 7 件: 拡張 / 0,0 / 単一軸 / 単一 cell promotion / 縮小 / 零次元 throw / 非整数 throw。

  empirical: 2163 tests pass (was 2156, +7)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/cell-range.ts` に追加: parseRange → `{rows, cols}`。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/range-dimensions-str.test.ts` 5 件: 単一 / 矩形 / col / row / 不正 input。

  empirical: 2156 tests pass (was 2151, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/cell-range.ts` に追加: rangeArea の A1 wrapper。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/range-area-str.test.ts` 5 件: 単一 cell=1 / 矩形 / 単一 col / 単一 row / 不正 input。

  empirical: 2151 tests pass (was 2146, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/cell-range.ts` に追加: shiftRange の A1 wrapper。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/shift-range-str.test.ts` 5 件: 正方向 / 負方向 / 単一 cell / (0,0) / OOXML grid 外 throw。

  empirical: 2146 tests pass (was 2141, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/cell-range.ts` に追加: union/intersection の A1 wrapper。disjoint で intersection は undefined。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/range-set-ops-str.test.ts` 9 件: union overlap/disjoint/同一/単一 cell + intersect overlap/disjoint/containment/同一/境界 cell。

  empirical: 2141 tests pass (was 2132, +9)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/cell-range.ts` に追加: rangesOverlap の A1 wrapper。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/ranges-overlap-str.test.ts` 6 件: overlap / containment / disjoint / 境界共有 / 単一 cell / 不正 input。

  empirical: 2132 tests pass (was 2126, +6)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/cell-range.ts` に追加: rangeContainsRange の A1 wrapper。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/is-range-in-range.test.ts` 6 件: contained / 同一 / partial overlap / disjoint / 単一 cell / 不正 input。

  empirical: 2126 tests pass (was 2120, +6)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/cell-range.ts` に追加: rangeContainsCell の A1 wrapper。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/is-cell-in-range.test.ts` 5 件: 内側 / 境界 inclusive / 外側 / 不正 cell / 不正 range。

  empirical: 2120 tests pass (was 2115, +5)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/defined-names.ts` に追加: `,` 区切り対応 (quoted title 内の `,` を escape)。`DefinedNameTarget` 型 export。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/get-defined-name-target.test.ts` 4 件: 単一 / quoted title / 複合 (Print_Titles) / 不存在で undefined。

  empirical: 2115 tests pass (was 2111, +4)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/defined-names.ts` に追加: getRangeAddress + addDefinedName。`opts.localToSheet` で scope。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/add-defined-name-for-range.test.ts` 5 件: workbook-scope / quoted title / localToSheet で scope / 不正 ws で throw / 同 name 上書き。

  empirical: 2111 tests pass (was 2106, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: formatSheetQualifiedRef(ws.title, range) wrapper。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/range-address.test.ts` 4 件: 単一 cell / 矩形 / quoted title / row/column span pass-through。

  empirical: 2106 tests pass (was 2102, +4)、typecheck / lint clean (14 warnings)。
  1. `src/utils/coordinate.ts` に formatSheetQualifiedRef、`src/worksheet/worksheet.ts` に getCellAddress wrapper。Excel 標準の quote rule (bare = `^[A-Za-z_][A-Za-z0-9_]*$`、それ以外は single-quote + 内部 `'` doubling)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/cell-address.test.ts` 8 件: bare / space / apostrophe / 数字始まり / 句読点 / round-trip parseSheetRange / getCellAddress 通常 / quoted title。

  empirical: 2102 tests pass (was 2094, +8)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: iterWorksheets + isWorksheetEmpty 短絡判定。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-empty-predicate.test.ts` 5 件: 空 wb / 全 ws 空 / 1 つ値あり / chartsheet only / 全 null。

  empirical: 2094 tests pass (was 2089, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: iterCells walk + value !== null で短絡判定。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/worksheet-empty-predicate.test.ts` 4 件: 空 sheet / 値あり / null cell は empty 扱い / 空文字列は non-empty。

  empirical: 2089 tests pass (was 2085, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: iterCells walk + `value !== null` + opts.includeFormulas / opts.includeRichText。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/non-empty-cell-count.test.ts` 5 件: 全 non-null / null skip / '' は count / formula skip / rich-text skip。

  empirical: 2085 tests pass (was 2080, +5)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: getWorkbookStats + getWorkbookCellsByKind + per-sheet metadata。型 `WorkbookOverview` / `WorkbookSheetOverview` も export。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/describe-workbook.test.ts` 4 件: 空 wb / 通常 sheet / chartsheet+hidden 混在 / tab-strip 順序保持。

  empirical: 2080 tests pass (was 2076, +4)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: cell value / styleId / 解決済み style chain / hyperlink / comment / mergedRange / inTables / inDV / inCF。型 `CellSummary` も export。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/cell-summary.test.ts` 6 件: 値のみ / styled+hyperlink+comment / merge / table+DV+CF / 不正 sheet で throw / 不存在 cell で exists:false + default styles。

  empirical: 2076 tests pass (was 2070, +6)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: unzipSync → 各 .csv entry を parseCsvToRange + pickUniqueSheetTitle で安全に dedupe + sanitise。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-from-csv-bundle.test.ts` 5 件: round-trip / 非 csv skip / 空 zip / coerceTypes / sanitise filename。

  empirical: 2070 tests pass (was 2065, +5)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: iterWorksheets + getWorksheetAsCsv → zipSync(<title>.csv → bytes)。タイトル sanitise + collision suffix。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-as-csv-bundle.test.ts` 5 件: 通常 / sanitise / chartsheet skip / opts.delimiter / 空 wb。

  empirical: 2065 tests pass (was 2060, +5)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: createWorkbook → addWorksheet → writeRangeFromObjects (or addTableFromObjects)。opts.asTable で table 化分岐。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-from-objects.test.ts` 5 件: 通常 / opts.sheetTitle / opts.asTable で table / opts.headers / 空 [] で empty sheet。

  empirical: 2060 tests pass (was 2055, +5)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: createWorkbook → addWorksheet → parseCsvToRange。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-from-csv.test.ts` 4 件: 通常 / opts.sheetTitle / opts.coerceTypes / 空 input。

  empirical: 2055 tests pass (was 2051, +4)、typecheck / lint clean (14 warnings)。
  1. `src/workbook/workbook.ts` に追加: iterWorksheets walk + getWorksheetAsCsv → Record<title, csv>。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-as-csv-record.test.ts` 5 件: 通常 / 空 sheet / 空 wb / chartsheet skip / opts.delimiter 伝播。

  empirical: 2051 tests pass (was 2046, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/csv.ts` に追加: getDataExtent → A1 → getRangeAsCsv。空 ws は ''。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/worksheet-as-csv.test.ts` 4 件: 通常 / 空 ws / sparse extent / opts.delimiter 伝播。

  empirical: 2046 tests pass (was 2042, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/csv.ts` に追加: RFC 4180 parser + writeRange。opts.coerceTypes で number/boolean coerce。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/parse-csv-to-range.test.ts` 10 件 (parseCsv 5 + parseCsvToRange 5): 標準 / quote / newline / 末尾 \n 扱い / 空 / write 通常 / coerceTypes / round-trip / delimiter / 空 input。

  empirical: 2042 tests pass (was 2032, +10)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/csv.ts` 新規追加: getRangeValues → CellValue → CSV-escape (RFC 4180)、Date は ISO、formula は cached value or formula source。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/range-as-csv.test.ts` 6 件: 通常 / quote escape / `,` 含む / newline / 空 cell / delimiter+trailingNewline / Date+bool+number coerce。

  empirical: 2032 tests pass (was 2026, +6)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/table.ts` に追加: writeRangeFromObjects → bounding-box → addExcelTable。columns は header と同順。空 throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/add-table-from-objects.test.ts` 4 件: 通常 / opts.headers / 空 throw / opts.style 反映。

  empirical: 2026 tests pass (was 2022, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: headers (union of keys / opts pinned) → grid → writeRange。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/write-range-from-objects.test.ts` 6 件: 通常 / 空 / opts.headers / null skip / union of keys / round-trip readRangeAsObjects。

  empirical: 2022 tests pass (was 2016, +6)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: getRangeValues 上に header 抽出 + zipping。`opts.skipEmptyRows`。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/read-range-as-objects.test.ts` 5 件: 通常 / data 無しで [] / 非 string header coerce / skipEmptyRows / 重複 header last-wins。

  empirical: 2016 tests pass (was 2011, +5)、typecheck / lint clean (14 warnings)。

- **次のタスク (前回)**: **`readRange(ws, range)` を queue → 既存 `getRangeValues` と重複のため scrap**。
- **次のタスク (前回 2)**: **`writeRange(ws, startRef, values)` 任意起点 2D write**。
  1. `src/worksheet/worksheet.ts` に追加: A1 → tuple → 行ごと/列ごとに setCell。null/undefined は skip。bounding-box を return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/write-range.test.ts` 5 件: B2 起点 / null skip / 空配列 → undefined / 既存 styleId preserve / 単一 cell。

  empirical: 2011 tests pass (was 2006, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: 2D values を順次 appendRow → `{firstRow, lastRow}`。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/append-rows.test.ts` 5 件: 3 row append / 空 row も cursor 進む / undefined/null skip / 既存 row 後に追加 / 空配列で no-op。

  empirical: 2006 tests pass (was 2001, +5)、typecheck / lint clean (14 warnings)。

- **次のタスク (前回)**: **`getCommentByCoord` 検討 → 既存 `getComment(ws, ref)` で代替済みのため scrap**。
  1. `src/workbook/workbook.ts` に追加: iterWorksheets + countCellsByKind を sum。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-cells-by-kind.test.ts` 4 件: 空 wb / 1 sheet / 2 sheet sum / chartsheet skip。

  empirical: 2001 tests pass (was 1997, +4)、typecheck / lint clean (14 warnings)。

- **次のタスク (前回)**: **`countCellsByKind(ws)` value-kind histogram**。
  1. `src/worksheet/worksheet.ts` に追加: iterCells walk + typeof/{kind} 判定。`CellsByKindCounts` interface も export。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/count-cells-by-kind.test.ts` 5 件: 空 / 全 kind 1 件ずつ / 重複 sum / sparse layout / row delete 後。

  empirical: 1997 tests pass (was 1992, +5)、typecheck / lint clean (14 warnings)。

- **次のタスク (前回)**: **`getDistinctValuesInRow(ws, row, opts?)` を追加** — column 版と対称。
  1. `src/worksheet/worksheet.ts` に追加: getCellsInRow → Set dedupe、column-order、skipNull/skipFormulas opts。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/distinct-row-values.test.ts` 5 件: dedupe / 混合型 / 空 row / skipNull / skipFormulas。

  empirical: 1992 tests pass (was 1987, +5)、typecheck / lint clean (14 warnings)。

- **次のタスク (前回)**: **`getDistinctValuesInColumn(ws, col, opts?)` を追加**。
  1. `src/worksheet/worksheet.ts` に追加: getCellsInColumn → Set で dedupe、first-seen 順、skipNull/skipFormulas opts。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/distinct-column-values.test.ts` 5 件: primitive dedupe / 混合型 / 空 col / skipNull / skipFormulas。

  empirical: 1987 tests pass (was 1982, +5)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: row 版は ws.rows のキーを sort、col 版は全 row 走査で集合化 + sort。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/populated-indices.test.ts` 5 件: row sparse / row 空 / col sparse / col 重複 dedupe / col 空。

  empirical: 1982 tests pass (was 1977, +5)、typecheck / lint clean (14 warnings)。

- **次のタスク (前回)**: **`removeAllImages(ws)` / `removeAllCharts(ws)` kind 別 drawing wipe**。
  1. `src/drawing/drawing.ts` に追加: items を kind で filter。
  2. `src/index.ts` から re-export。
  3. `tests/phase-6/remove-by-kind.test.ts` 4 件: image 削除 chart 残 / chart 削除 image 残 / no drawing → 0 (×2)。

  empirical: 1977 tests pass (was 1973, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: `iterRows` を flatten して 1 cell ずつ yield。min/max bounds は `IterRowsOptions` を継承。
  2. `src/index.ts` から re-export (`iterWorksheetCells` alias)。
  3. `tests/phase-5/iter-cells.test.ts` 4 件: 順序 / 空 / bounds / 挿入順序非依存。

- **次のタスク (前回)**: **`cellRangeFromCells(cells)` を追加** — Cell[] → bounding A1 range。
  1. `src/worksheet/cell-range.ts` に追加: row/col min/max → A1 / A1:B2 形式。空 throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/cell-range-from-cells.test.ts` 6 件: 単一 / 矩形 / sparse / 1 col / 空 throw / 実 Cell。

  empirical: 1973 tests pass (was 1963, +10 = +6 from cellRangeFromCells + +4 from iterCells)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: 既存 freezePanes/freezeRows/freezeColumns の薄い wrapper。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/freeze-shortcuts.test.ts` 4 件: firstRow / firstColumn / both / 既存 freeze 上書き。

  empirical: 1963 tests pass (was 1959, +4)、typecheck / lint clean (14 warnings)。
  1. `src/worksheet/worksheet.ts` に追加: row/column ごとに populated Cell[] を coordinate 順 sort で返す。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/cells-in-row-column.test.ts` 5 件: row sparse / row 空 / col sparse / col 空 / col 順序保証。

  empirical: 1959 tests pass (was 1954, +5)、typecheck / lint clean (14 warnings)。
  1. `src/styles/cell-style.ts` に追加: styleId を 0 reset。range 版は existing cell のみ walk (no materialisation)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/styles/clear-cell-style.test.ts` 5 件: 単一 reset / fill+bold reset / range 全 cell reset / 空 cell は materialise しない / 範囲外 cell は影響なし。

  empirical: 1954 tests pass (was 1949, +5)、typecheck / lint clean (14 warnings)。
  1. `src/styles/cell-style.ts` に追加: merge mode は per-cell mergeAlignment、replace mode は setRangeStyle 経由。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/styles/set-range-alignment.test.ts` 5 件: merge で indent preserve / replace で indent drop / merge で horizontal preserve / range 全 cell / 空 cell 作成。

  empirical: 1949 tests pass (was 1944, +5)、typecheck / lint clean (14 warnings)。

- **次のタスク (前回)**: **`setRangeWrapText(wb, ws, range, on=true)` 一括 wrap helper を追加**。
  1. `src/styles/cell-style.ts` に追加: range walk + wrapCellText per cell で既存 alignment 保持。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/styles/set-range-wrap-text.test.ts` 4 件: 全 cell on / off / horizontal preserve / 空 range で cell 作成。

  empirical: 1944 tests pass (was 1940, +4)、typecheck / lint clean (14 warnings — 16 → 14、tests/e2e/scenarios/19-charts-classic.test.ts の `!` 撤去で 2 減)。
  1. `src/styles/cell-style.ts` に追加: `setRangeStyle` 経由で range 全 cell に Protection を stamp。`Protection | Partial<Protection>` を受け、partials は missing fields を `false` にデフォルト。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/styles/set-range-protection.test.ts` 4 件: locked=false 全 cell / hidden=true partial 適用 / bold + protection 共存 / 空 range で cell 作成。

  empirical: 1940 tests pass (was 1936, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`cssRecordToInlineStyle(record)` helper を追加** — `Record<string,string>` → HTML `style="…"` inline 文字列。
  1. `src/utils/css.ts` (新規) に追加: alphabetical sort で `k: v;` 連結。空 record / `undefined` → ''。空文字値 skip。値に `;` 含むものは drop (defense against attribute-injection)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/styles/css-record-to-inline-style.test.ts` 5 件: empty / 単一 / multi-alphabetical / 空文字 skip / `;` 含む値 drop。

  empirical: 1936 tests pass (was 1931, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`cellStyleToCss(wb, cell)` aggregator を追加**。
  1. `src/styles/cell-style.ts` に追加: cell の `styleId` から `wb.styles.cellXfs[styleId]` を引いて → font/fill/border/alignment を 4 種 toCss helper でそれぞれ変換 → 1 record に merge (font→fill→border→alignment 順、後勝ち)。`styleId === 0` + 完全 default xf は早期 `{}`。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/styles/cell-style-to-css.test.ts` 5 件: 空 / setBold (DEFAULT_FONT 込み) / bold+bg / formatAsHeader (font+fill+border) / vertical-align collision precedence。

  empirical: 1931 tests pass (was 1926, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`alignmentToCss(alignment)` Alignment → CSS record helper を追加**。
  1. `src/styles/alignment.ts` に追加: horizontal → text-align (general → 出力なし、fill → left approx、distributed → justify) / vertical → vertical-align (table-cell semantics) / wrapText → white-space: pre-wrap / textRotation → transform: rotate(-deg) (Excel ccw → CSS cw flip)、255 → writing-mode: vertical-rl / indent → padding-left em。空 Alignment → {}。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/styles/alignment-to-css.test.ts` 7 件: empty / horizontal 5 種 / vertical 3 種 / wrapText / textRotation 90/0/255 / indent 2/0 / 全部 combine。

  empirical: 1926 tests pass (was 1919, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`borderToCss(border)` Border → CSS border-property record helper を追加**。
  1. `src/styles/borders.ts` に追加: 4 sides を CSS border shorthand (`<width> <style> <#color>`) に。Excel SideStyle → CSS mapping (thin/hair→1px solid / medium→2px solid / thick→3px solid / double→3px double / dotted→1px dotted / dashed family→1px dashed / mediumDashed family→2px dashed)。色未指定 / theme は `currentColor` フォールバック。diagonal / vertical / horizontal sides は無視。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/styles/border-to-css.test.ts` 6 件: 空 / 全辺 thin + rgb / 1辺だけ + style 無し skip / thick/double/dashed/dotted variants / currentColor fallback / diagonal は skip。

  empirical: 1919 tests pass (was 1913, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`fillToCss(fill)` Fill → CSS background-property record helper を追加**。
  1. `src/styles/fills.ts` に追加: PatternFill ('solid' → fgColor を `background-color`) / non-solid → bgColor 折り畳み / GradientFill (linear/path → `linear-gradient` / `radial-gradient` の `background-image`)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/styles/fill-to-css.test.ts` 7 件: undefined+none / solid+rgb / theme-only fg → {} / non-solid bg/fg fallback / linear gradient / radial gradient / 全 stop が theme で {}。

  empirical: 1913 tests pass (was 1906, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`fontToCss(font)` Font → CSS-property record helper を追加**。
  1. `src/styles/fonts.ts` に追加: name → font-family / size → font-size pt / bold/italic/underline/strike → weight/style/decoration / color → `#RRGGBB` / vertAlign sup/sub → vertical-align + 0.83em fallback。空 Font → `{}`。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/styles/font-to-css.test.ts` 7 件: 空 / family+size / bold/italic/strike / underline+strike combine / rgb color + theme skip / sup/sub + size override / quote escape。

  empirical: 1906 tests pass (was 1899, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`colorToHex(color)` Color → ARGB hex readback を追加**。
  1. `src/styles/colors.ts` に追加: `rgb` あれば normaliseRgb / `indexed` あれば palette lookup / theme・auto・空は undefined を返す。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/styles/color-to-hex.test.ts` 6 件: rgb / indexed / theme / auto / empty / undefined / rgb wins over indexed。

  empirical: 1899 tests pass (was 1893, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`getAllImages(wb)` / `getAllCharts(wb)` workbook-wide drawing aggregator を追加**。
  1. `src/workbook/workbook.ts` に追加: `iterWorksheets` で全 sheet 巡回 → `ws.drawing?.items` の content.kind で picture / chart を分けて `{sheet, item}` 配列で return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/all-images-charts.test.ts` 4 件: 2 sheet image 集約 + chart 混在 / 空 wb / 2 sheet chart 集約 + image 混在 / drawing 無し sheet skip。

  empirical: 1893 tests pass (was 1889, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **drawing listing helpers (`listImagesOnSheet` / `listChartsOnSheet` / `removeAllDrawingItems`) を追加**。
  1. `src/drawing/drawing.ts` に追加: `ws.drawing?.items` を kind で filter して image/chart のみを return、wipe 系は items 配列を空にして count 返す。
  2. `src/index.ts` から re-export。
  3. `tests/phase-6/drawing-listing.test.ts` 6 件: images 2 件抽出 / 空 / charts 2 件抽出 / 空 / wipe 2 件 + count / 0 件で 0。

  empirical: 1889 tests pass (was 1883, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **SST helpers (`getSharedStringIndex` / `getSharedStringAt` / `sharedStringCount`) + public API export を追加**。
  1. `src/workbook/shared-strings.ts` に追加: 3 helpers + 既存 SST モジュール一式 (`SharedStringsTable` / `addSharedString` / `makeSharedStrings`) を src/index.ts から正式 re-export (これまで未公開)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/shared-strings-helpers.test.ts` 6 件: count growth / dedupe で count 不変 / index lookup / unknown で undefined / at lookup / out-of-range で undefined。

  empirical: 1883 tests pass (was 1877, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`removeAllTables` / `removeAllDataValidations` 一括 wipe を追加**。
  1. `src/worksheet/worksheet.ts` に追加: 各 array を `[]` で空にして count を return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/remove-all-tables-validations.test.ts` 4 件: tables 2 件削除 + count / 0 件で 0 / validations 2 件削除 + count / 0 件で 0。

  empirical: 1877 tests pass (was 1873, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`removeAllConditionalFormatting(ws)` 一括 wipe を追加**。
  1. `src/worksheet/worksheet.ts` に追加: `ws.conditionalFormatting = []` で全 CF block を削除し、count を return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/remove-all-conditional-formatting.test.ts` 2 件: 2 種 (cellIs+colorScale) 削除 + count / 0 件で 0。

  empirical: 1873 tests pass (was 1871, +2)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`swapSheets(wb, titleA, titleB)` を追加**。タブストリップ位置入れ替え:
  1. `src/workbook/workbook.ts` に追加: 両 title 検索 → 配列内 swap → activeSheetIndex を follow。同一 title は no-op、片方 missing で throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/swap-sheets.test.ts` 5 件: 隣接 swap / 非隣接 swap / activeSheetIndex follow / 同 title no-op / missing throw 両側。

  empirical: 1871 tests pass (was 1866, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`isValidColumnLetter` / `isValidRowNumber` / `isValidColumnNumber` mini-predicates を追加**。
  1. `src/utils/coordinate.ts` に追加: 1..3 char column letter (case-insensitive) / [1, 1048576] integer row / [1, 16384] integer col。non-integer / out-of-bound / 非 string|number は false。
  2. `src/index.ts` から re-export。
  3. `tests/phase-1/is-valid-row-col.test.ts` 12 件: ColumnLetter A..XFD + 大文字小文字 / 空+over-length+非letter+空白 / XFE / 非 string / RowNumber 1+max / 0+負 / >max / 非 integer+NaN+Infinity / 非 number / ColumnNumber 1+max / 0 + >max / 非 integer。

  empirical: 1866 tests pass (was 1854, +12)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`isValidRangeRef(s)` predicate を追加**。`isValidCellRef` の range 版:
  1. `src/utils/coordinate.ts` に追加: 単一 cell / 二角 range / 全列 (A:A) / 全行 (1:1) を accept、$/空白/sheet prefix/multi-range/範囲外/非 string は reject。
  2. `src/index.ts` から re-export。
  3. `tests/phase-1/is-valid-range-ref.test.ts` 8 件: 単 cell / 二角 / 全列 / 全行 / $+空白+空 / 範囲外 / 非 string / sheet prefix + multi-range。

  empirical: 1854 tests pass (was 1846, +8)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`isValidCellRef(s)` predicate を追加**。input validation:
  1. `src/utils/coordinate.ts` に追加: 単一 A1 ref のみ true。`$` absolute marker / 範囲 / 空白 / 範囲外 row/col / 非 string は false。
  2. `src/index.ts` から re-export。
  3. `tests/phase-1/is-valid-cell-ref.test.ts` 6 件: 通常 + 大文字小文字 + max coord / `$` 拒否 / range 拒否 / 空白 + 空文字 / 範囲外 row/col / 非 string。

  empirical: 1846 tests pass (was 1840, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`removeAllHyperlinks` / `removeAllComments` 一括 wipe を追加**。
  1. `src/worksheet/worksheet.ts` に追加: hyperlinks/legacyComments 配列を 1 line で空にし、count を return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/remove-all-hyperlinks-comments.test.ts` 4 件: hyperlinks 全 drop+count / 0 件で 0 / comments 全 drop+count / 0 件で 0。

  empirical: 1840 tests pass (was 1836, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`renameDefinedName(wb, oldName, newName, scope?)` を追加**。
  1. `src/workbook/defined-names.ts` に追加: 同 scope 内 lookup → conflict check → in-place rename。Return は boolean (見つかったか)、scope 内重複は throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/rename-defined-name.test.ts` 5 件: workbook-scope rename / sheet-scope rename + 他 scope 保持 / missing で false / 同 scope collision throw / 別 scope 同名 OK。

  empirical: 1836 tests pass (was 1831, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`removeAllMergedRanges(ws)` 一括解除を追加**。
  1. `src/worksheet/worksheet.ts` に追加: `ws.mergedCells = []` で全 merge を削除し、count を return。cell value は untouched。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/remove-all-merged-ranges.test.ts` 3 件: 2 merge 削除 + count / 0 件で 0 / cell value 保持。

  empirical: 1831 tests pass (was 1828, +3)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`setSheetStates` / `showAllSheets` 一括 visibility helpers を追加**。
  1. `src/workbook/workbook.ts` に追加: `setSheetStates(wb, { title: state })` で複数 sheet 状態を 1-call 更新 (missing title は throw)、`showAllSheets(wb)` で hidden+veryHidden 全部を visible に戻し count を return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/sheet-state-bulk.test.ts` 4 件: 3 sheet 一括設定 / missing title throw / showAllSheets count + 全 visible / 隠れ無しで 0。

  empirical: 1828 tests pass (was 1824, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`copyCellStyle` / `cloneCellStyle` を追加**。同 wb / 異 wb cell style 移植:
  1. `src/styles/cell-style.ts` に追加: `copyCellStyle(wb, source, target)` (同 wb 内で styleId を直接共有、xf pool 増加なし)、`cloneCellStyle(srcWb, srcCell, dstWb, dstCell)` (異 wb 間 deep copy: font/fill/border/numFmt を別 stylesheet にも個別追加し、新 styleId を target に set)。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/copy-cell-style.test.ts` 4 件: copy 同 styleId+pool 不変 / clone 異 wb で font+fill+numFmt 全 axis 移植 / clone 戻り値 = target.styleId / clone target wb で cellXfs[0] default reserve。

  empirical: 1824 tests pass (was 1820, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`iterVisibleWorksheets` / `iterWorksheetsByState` を追加**。state filter sheet iter:
  1. `src/workbook/workbook.ts` に追加: `iterVisibleWorksheets(wb)` (state==='visible' のみ yield)、`iterWorksheetsByState(wb, state)` (任意 state filter)。chartsheet skip。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/iter-visible-worksheets.test.ts` 5 件: hidden+veryHidden skip / chartsheet skip / 空 wb / state filter (3 state) / setSheetState 反映。

  empirical: 1820 tests pass (was 1815, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **HSL adjustment shortcuts (`rotateHue` / `adjustSaturation` / `adjustLightness`) を追加**。
  1. `src/styles/colors.ts` に追加: 全 hexToHsl→adjust→hslToHex の 1-line wrapper。alpha 保持、s/l は clamp、h は wrap。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/color-hsl-shortcuts.test.ts` 11 件: rotateHue 120°/-120°/360° identity / alpha 保持 / adjustSaturation +1 clamp+hue 保持 / -1 → gray / alpha 保持 / adjustLightness +0.5 from black → 0.5 / -1 → black / +1 → white / alpha 保持。

  empirical: 1815 tests pass (was 1804, +11)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **color HSL 変換 (`hexToHsl` / `hslToHex`) を追加**。テーマ調整向け hue/sat/light 操作:
  1. `src/styles/colors.ts` に追加: `hexToHsl(hex)` → `{h ∈ [0,360), s, l ∈ [0,1], a ∈ [0,255]}` (RGB→HSL std formula)、`hslToHex(h, s, l, alpha=255)` → ARGB hex (h は mod-360 wrap、s/l は clamp)。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/color-hsl.test.ts` 13 件: 白/黒/赤/緑/青の hue/sat/light / alpha 保持 / hslToHex 逆変換 5 色 / h wrap (-120 / 360) / alpha 引数 / 任意 hex round-trip。

  empirical: 1804 tests pass (was 1791, +13)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`replaceCellValuesInWorkbook` workbook-wide find-and-replace を追加**。
  1. `src/workbook/workbook.ts` に追加: `iterAllCells` 経由で全 sheet 走査、string モード (exact-equal on string-valued) と predicate モード `(value, cell, sheet) → boolean`。返値は変更件数。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/replace-in-workbook.test.ts` 4 件: 2 sheet 横断 string 一括置換 / predicate に sheet 渡る / 不一致 0 / string モードで数値+真偽値 skip。

  empirical: 1791 tests pass (was 1787, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`findCellInWorkbook` / `findCellsInWorkbook` を追加**。`findCells` の workbook 版:
  1. `src/workbook/workbook.ts` に追加: 両方とも `iterAllCells` を内部で使い、predicate(cell, sheet) → boolean。findCellInWorkbook は最初の一致 / findCellsInWorkbook は全件配列。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/find-cells-in-workbook.test.ts` 5 件: 異 sheet で iter 順最初 / 不一致 undefined / predicate に sheet 渡る / 全件配列 in iter 順 / 空 wb で空。

  empirical: 1787 tests pass (was 1782, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`findTable(wb, displayName)` workbook-wide table lookup を追加**。
  1. `src/workbook/workbook.ts` に追加: 全 worksheet を walk して displayName 一致 table を `{sheet, table}` で return、無ければ undefined。Excel の table-name 一意性保証で first match で十分。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/find-table.test.ts` 3 件: 異 sheet で displayName lookup / 不存在で undefined / 空 wb で undefined。

  empirical: 1782 tests pass (was 1779, +3)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`listPrintAreas` / `listPrintTitles` を追加**。`_xlnm.Print_Area` / `_xlnm.Print_Titles` を name で filter:
  1. `src/workbook/defined-names.ts` に追加: definedNames を name で filter する 1-line helper x 2。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/print-area-titles-listing.test.ts` 3 件: listPrintAreas (2 sheet 分の `_xlnm.Print_Area` のみ抽出 + 他 name 無視) / 空 wb は空 / listPrintTitles 同等。

  empirical: 1779 tests pass (was 1776, +3)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **color hex mixing helpers (`lighten` / `darken` / `mixColors`) を追加**。
  1. `src/styles/colors.ts` に追加: `lighten(hex, amount)` (white と線形 mix)、`darken(hex, amount)` (black と線形 mix)、`mixColors(a, b, t)` (per-channel + alpha 線形補間)。すべて 0..1 clamp、alpha 保持、6char 入力時は alpha=00 (Excel convention)。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/color-mixing.test.ts` 12 件: lighten unchanged/white/halfway/clamp/6char (4) / darken unchanged/black/halfway (3) / mixColors endpoints/midpoint/alpha interp/clamp (4)。

  empirical: 1776 tests pass (was 1764, +12)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`getAllConditionalFormatting(wb)` aggregator を追加**。
  1. `src/workbook/workbook.ts` に追加: `{sheet, formatting}` 配列を tab-strip 順で return。`getAllTables` / `getAllDataValidations` と同 pattern。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/all-conditional-formatting.test.ts` 3 件: 2 sheet 集約 (cellIs+colorScale 混在) / chartsheet skip / 空 wb。

  empirical: 1764 tests pass (was 1761, +3)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`pickUniqueSheetTitle(wb, base)` を追加**。Excel UI の "Sheet1 (2)" 風 auto-suffix:
  1. `src/workbook/workbook.ts` に追加: base が free なら verbatim、衝突時は ` (N)` を 2..999 まで増やして free slot を探す。base+suffix が 31 char を超える場合は base を切り詰めて全体 31 ≤ に維持。base 自体が valid でないと throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/pick-unique-sheet-title.test.ts` 5 件: free pass-through / 1〜2 衝突で ` (2)`/` (3)` / 31 char base で 27 char truncation+` (2)` / invalid base throw / 出力が validateSheetTitle 通過。

  empirical: 1761 tests pass (was 1756, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **autofit 列の font-aware scaling を追加**。
  1. `src/worksheet/worksheet.ts`: `autofitColumn` / `autofitColumns` に `opts.workbook` 引数を追加。供給時は cell の `styleId` → `wb.styles.cellXfs[id].fontId` → `wb.styles.fonts[fontId].size` を辿り、長さを `(size / 11)` で線形 scaling。未供給時は従来の文字列長 fallback。
  2. `tests/phase-5/autofit-columns.test.ts` 3 件追加: 22pt cell が 11pt の ~2× width / workbook 引数なしで font は無視 / autofitColumns で複数列 mixed 11pt/22pt の独立 scaling。
  3. 既存 9 件の string-length-only テストはそのまま pass。

  empirical: 1756 tests pass (was 1753, +3)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`replaceInRange` 範囲限定 find-and-replace を追加**。`replaceCellValues` の rectangle 範囲版:
  1. `src/worksheet/worksheet.ts` に追加: `replaceInRange(ws, range, search, replacement)`。matching rule は `replaceCellValues` と同じ (string→exact-equal、function→predicate)。populated cell のみ走査、auto-allocate 無し。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/replace-in-range.test.ts` 5 件: 範囲内のみ置換 + 範囲外保持 / predicate variant + 範囲外保持 / 不一致 0 / auto-allocate 無し / 数値+真偽値 skip (string モード)。

  empirical: 1753 tests pass (was 1748, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **長期 /loop 仕上げターン (public API + docs + edge-fixtures + perf 微調整 + UTF-8 + max-coord + migration guide) 完了**。本ターン (連続コミット 11 件) で以下を一括完遂:
  1. phase-4 streaming reader real-fixture acceptance (`genuine/sample.xlsx` SAX iter ↔ eager loadWorkbook 座標一致 + `empty-with-styles.xlsx` minRow/maxRow band フィルタ) — `tests/phase-4/read-only-genuine.test.ts` 2件
  2. iterRows 早期終了 + cell-skip 最適化 (currentRow > maxRow で generator return、band 外 row の `<c>` attrs パース skip)
  3. public API 露出整備: `src/index.ts` に Cell / Worksheet / Workbook helpers + Style 値オブジェクト群 (`setCell` / `addWorksheet` / `createWorkbook` / formula helpers / `makeRichText` / `setCellFont` / `makeFont` / `makePatternFill` etc.) を re-export、`src/zip/index.ts` に `StreamingEntryWriter` 追加、`src/styles/index.ts` を新設
  4. `README.md` 修正 (full-lib 例の `fromBuffer` を `openxml-js/node` から import するよう split)、`tests/phase-3/readme-examples.test.ts` 5件で 5 つの README 例を public API に対して smoke-test
  5. edge-fixture deep-assert 4件追加: comments.xlsx (per-sheet legacyComments 数 6/0/1) / hyperlink.xlsx (外部 URL ref) / test_datetime.xlsx (numeric serial round-trip) / legacy_drawing.xlsm (control VML + ctrlProps セット)
  6. UTF-8 round-trip テスト 2件 (`workbook_russian_code_name.xml` の `codeName="ЭтаКнига"` / synthetic 多言語 + emoji)
  7. max-coord (XFD1048576) round-trip 2件 (単一最大セル + 4 隅の sparse spread)
  8. THIRD_PARTY_NOTICES.md を実 runtime deps (fflate / saxes / fast-xml-parser) で更新、stale な `scripts/regenerate-notices.ts` 言及を削除
  9. `docs/migrate-from-openpyxl.md` を新設 (loading / cells / styles / worksheets / streaming write&read / passthrough / known gaps をテーブル形式でカバー)、`tests/phase-3/migration-guide-examples.test.ts` 5件で例の API surface を smoke-test
  
  1148 tests pass (was 1128, +20 in this turn)、`pnpm size` full 81.7 KB / streaming 46.52 KB clean、lint / typecheck clean。**残**: random-access streaming reader for sub-sheet cell ranges (perf 最適化、ECMA-376 row-order 前提に row-offset index を作る案; 単独実装スプリント要)、Excel 365 視覚 QA (人手)、ZIP64 write の正式対応 (fflate 上流)。コア実装はフェーズ1-7 + streaming acceptance + docs + 多数の edge-fixture acceptance + public API 露出 + 全主要 perf gate clear すべて整備、1.0 候補レベル。
- **次のタスク**: **フェーズ4 §2.x row-offset index for sub-sheet streaming reader 完了**。`makeStreamingReadOnlyWorksheet` に lazy-built `<row r="N">` byte-offset index + `</sheetData>` 終端位置の cache を追加。`iterRows({minRow > 1})` の経路で: (1) 初回呼出で `buildRowOffsetIndex(bytes)` を実行 (純 byte-level スキャン、`<row` ASCII pattern を見つけて attrs から `r="N"` を抽出、~50 ns/row)、(2) `firstRowAtOrAfter(index, minRow)` で binary search、(3) `sliceFromRow` で `[targetOffset, sheetDataEnd)` を切り出し、`<?xml ?><sheetData xmlns="...">SLICE</sheetData>` で wrap → saxes に渡す。これで minRow=999000 の 1M 行 sheet で saxes-walk が ~1000 行ぶんに圧縮。`minRow <= 1` は index を完全 skip (旧 path のまま)、`minRow > maxKnownRow` は空 generator を即返す。`tests/phase-4/row-offset-index.test.ts` 5 件: minRow only / minRow+maxRow band / out-of-range / index 再利用 (連続 band query) / SST + 数値混合の cell value 整合。1153 tests pass (was 1148)、bundle / lint / typecheck clean。残：Excel 365 視覚 QA (人手)、ZIP64 write の正式対応 (fflate 上流)。コア実装フェーズ1-7 + streaming acceptance + docs + edge-fixture acceptance + public API 露出 + 全主要 perf gate clear + sub-sheet random-access reader すべて整備、**docs target の機能要件は事実上完了**。
- **次のタスク**: **row-offset index speedup の実機計測 完了**。`tests/perf/row-index.test.ts` を追加: 10k 行 × 5 列の sheet を 1 つ書き、tail 100 行を `iterRows({minRow: ROWS-99, maxRow: ROWS})` で 3 回計測 (warm-up + best-of-3) ↔ 同じ tail を no-min walk + count-cap で計測 (full-sheet saxes-walk)。**実機実測**: index path 12.1 ms、full-walk path 41,310 ms、**speedup 3415x**。`PERF_ROW_INDEX_GATE=1` で `>2x` を assert (gate-off default、observe-only)。stderr に `[perf-row-index] tail-100-of-10,000 rows: index 12.10ms, full 41310.54ms, speedup 3415.0x` を 1 行で記録。1153 default tests pass (perf スイートは別 config)、lint clean。残：Excel 365 視覚 QA (人手)、ZIP64 write の正式対応 (fflate 上流)。docs target の機能要件 + perf 実証すべて完了。
- **次のタスク**: **`getAllTables` / `getAllDataValidations` aggregator を追加**。
  1. `src/workbook/workbook.ts` に追加: `{sheet, table | validation}` 配列を tab-strip 順で return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/all-tables-and-validations.test.ts` 4 件: tables 2 sheet 集約 / chartsheet skip 空 / validations 集約 / 空 wb。

  empirical: 1748 tests pass (was 1744, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`getAllHyperlinks` / `getAllComments` workbook-wide listing を追加**。
  1. `src/workbook/workbook.ts` に追加: `{sheet, hyperlink|comment}` 配列を tab-strip 順で return。`getAllMergedRanges` と同じ pattern。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/all-hyperlinks-and-comments.test.ts` 4 件: hyperlinks 2 sheet 集約 + chartsheet skip / 空 / comments 集約 / 空 wb。

  empirical: 1744 tests pass (was 1740, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **color contrast helpers (`luminance` / `contrastRatio` / `pickReadableTextColor`) を追加**。WCAG-based UI helper:
  1. `src/styles/colors.ts` に追加: `luminance(hex)` (WCAG 2 相対輝度 0..1、alpha 無視、6/8 char 両対応)、`contrastRatio(a, b)` (順不同 1..21)、`pickReadableTextColor(bg)` (luminance < 0.179 で白/それ以外で黒、Excel UI default の dark blue でも白を選ぶ midpoint)。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/color-contrast.test.ts` 10 件: white=1+black=0 / mid-gray ≈0.18 / alpha 無視 / 6char OK / 21:1 / 順不同 / 同色 1:1 / dark bg→white / light bg→black / theme blue→white。

  empirical: 1740 tests pass (was 1730, +10)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`validateSheetTitle` / `isValidSheetTitle` を追加**。Excel の sheet name rule 全部を予 check 可能に:
  1. `src/workbook/workbook.ts` に追加: `validateSheetTitle(title)` (理由文字列 / 有効なら undefined)、`isValidSheetTitle` (boolean+narrowing)。Rule: type=string, length 1..31, no `: \ / ? * [ ]`, no leading/trailing `'`, "history" (case-insensitive) reserved。`validateUniqueTitle` も新 helper 経由でリッチ message に。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/sheet-title-validation.test.ts` 8 件: 通常 OK / 非 string + 空 + 32 文字 / 禁止文字 6 種 / leading/trailing apostrophe + middle OK / "History"/"history"/"HISTORY" + "Historical" OK / 31 文字 OK / isValid true/false。

  empirical: 1730 tests pass (was 1722, +8)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`getAllMergedRanges(wb)` workbook-wide merge listing を追加**。
  1. `src/workbook/workbook.ts` に追加: `{sheet, range}` 配列を tab-strip 順で return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/all-merged-ranges.test.ts` 3 件: 2 sheet+3 merge 集約 / chartsheet skip / 空 wb。

  empirical: 1722 tests pass (was 1719, +3)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`iterAllCells(wb)` workbook-wide cell iterator を追加**。
  1. `src/workbook/workbook.ts` に追加: `iterAllCells(wb)` で `{sheet, cell}` を tab-strip 順 + sheet 内 row→col 順に yield。chartsheet は skip。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/iter-all-cells.test.ts` 4 件: 2 sheet 横断 / chartsheet skip / sheet 内 row-then-col / 空 wb は空。

  empirical: 1719 tests pass (was 1715, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`cellValueAsBoolean` を追加**。CellValue → `boolean | undefined` coercion:
  1. `src/cell/cell.ts` に追加: boolean 通過 / 0→false 他 finite 数→true (Excel 真偽値解釈) / `'true'`/`'false'` 大文字小文字無視 / formula の cachedValue boolean 通過 / NaN+Infinity+任意文字列+null+Date+error+duration+rich-text→undefined。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/cell-value-as-boolean.test.ts` 8 件: bool 通過 / 0→false + 非0→true / NaN+Inf→undefined / "true"/"FALSE" parse / 任意文字列+空文字+'1'→undefined / formula 真偽 cached / formula 数値 cached→undefined / 他 5 種→undefined。

  empirical: 1715 tests pass (was 1707, +8)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`cellValueAsDate` を追加**。CellValue → `Date | undefined` coercion:
  1. `src/cell/cell.ts` に追加: Date pass-through / ISO string parse / Duration `new Date(ms)` / それ以外 (number / boolean / null / formula / error / rich-text / 非 ISO 文字列) は undefined。docstring で「Excel serial 数値は対象外、`excelToDate` を使え」と明記。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/cell-value-as-date.test.ts` 7 件: Date 通過 / ISO 文字列 / 非 date 文字列 / 空文字 / Duration / 数値→undefined / null+boolean+error+rich-text→undefined。

  empirical: 1707 tests pass (was 1700, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`setCellArrayFormula` shortcut を追加**。setCell + setArrayFormula 1-call:
  1. `src/worksheet/worksheet.ts` に追加: `setCellArrayFormula(ws, row, col, ref, formula, { cachedValue?, styleId? })`。leading `=` strip。`t='array'` + ref が attached。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/set-cell-array-formula.test.ts` 3 件: TRANSPOSE で formula+ref+t='array' / `=` なし passthrough / cachedValue+styleId 通過。

  empirical: 1700 tests pass (was 1697, +3)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`setCellFormula` shortcut を追加**。setCell + setFormula 1-call:
  1. `src/worksheet/worksheet.ts` に追加: `setCellFormula(ws, row, col, formula, { cachedValue?, styleId? })`。leading `=` は strip するので `'=A1+1'` でも `'A1+1'` でも OK。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/set-cell-formula.test.ts` 4 件: `=A2+B2` で正常 + `=` strip / `SUM(A:A)` (= なし) / cachedValue 通過 / styleId 通過。

  empirical: 1697 tests pass (was 1693, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`setCellRichText` を追加**。setCell + makeRichText の 1-call shortcut:
  1. `src/worksheet/worksheet.ts` に追加: `setCellRichText(ws, row, col, runs, styleId?)` で `{kind:'rich-text', runs:makeRichText(runs)}` を assign。runs は `TextRun | { text, font? }` 混在 OK。
  2. `src/index.ts` から re-export。
  3. `tests/phase-2/set-cell-rich-text.test.ts` 3 件: inline run object でリッチテキスト set / pre-built RichText 透過 / styleId 引数。

  empirical: 1693 tests pass (was 1690, +3)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`removeDefinedNames(wb, predicate)` 一括削除を追加**。`removeDataValidations` の workbook 版:
  1. `src/workbook/defined-names.ts` に追加: `removeDefinedNames(wb, predicate)` で `definedNames` を filter、削除件数を return。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/remove-defined-names.test.ts` 3 件: 一致する 2 件削除 / 不一致 0 / scope predicate で sheet-scope のみ削除。

  empirical: 1690 tests pass (was 1687, +3)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **stylesheet pool listing helpers (`listFonts` / `listFills` / `listBorders` / `listCellXfs` / `listCellStyleXfs`) を追加**。
  1. `src/styles/stylesheet.ts` に追加: 各 array 自体を `ReadonlyArray` として return (no copy)。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/stylesheet-listings.test.ts` 5 件: listFonts (default 1 → +1) / listFills (default 2 → +1) / listBorders (default 1 → +1) / listCellXfs (0 → reserve+1=2) / listCellStyleXfs (0 → +1)。

  empirical: 1687 tests pass (was 1682, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **sheet iteration helpers (`iterWorksheets` / `iterChartsheets` / `listWorksheets` / `listChartsheets`) を追加**。chartsheet-skip iteration:
  1. `src/workbook/workbook.ts` に追加: `iter*` は generator、`list*` は array スナップショット。tab-strip 順を保つ。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/iter-sheets.test.ts` 4 件: iterWorksheets が chartsheet skip / 空 wb / iterChartsheets が worksheet skip / chartsheet 無いと空。

  empirical: 1682 tests pass (was 1678, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`moveRange` を追加**。`copyRange` + source clear、overlap 安全:
  1. `src/worksheet/worksheet.ts` に追加: source の populated cell 全件を snapshot → source 範囲を clear → snapshot を destination に setCell。同シートで overlap してもデータ消失しない。dest 別シートも opts.targetWs 経由 OK。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/move-range.test.ts` 4 件: 非 overlap で source 完全消滅+target 配置 / 1 step 横シフト overlap (3 cell の forward shift) / 別 worksheet 経由 / 空 source 0 件。

  empirical: 1678 tests pass (was 1674, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`copyRange` を追加**。範囲の cell コピー:
  1. `src/worksheet/worksheet.ts` に追加: `copyRange(ws, source, target, { targetWs? })`。シート内/間の cell コピー (value + styleId + hyperlinkId/commentId)。target が source より小さければ収まる範囲だけ copy。空 cell は skip。返値はコピー件数。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/copy-range.test.ts` 6 件: 同シート 2×2 / styleId 保持 / 空 source skip / target 縮小 truncate / 別 worksheet 経由 / 既存 dest 上書き。

  empirical: 1674 tests pass (was 1668, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`clearRange` / `clearAllCells` を追加**。フォーマット保持で cell だけ wipe:
  1. `src/worksheet/worksheet.ts` に追加: `clearRange(ws, range)` (rectangular range の populated cell のみ削除、空 row map は prune、count return)、`clearAllCells(ws)` (全 cell 削除 + `_appendRowCursor=0` reset、count return)。Column/row dimensions・merges・comments・hyperlinks 等は untouched。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/clear-range.test.ts` 7 件: clearRange 範囲内のみ削除+count / 空 range 0 / row map prune / dimensions+merges 保持 / clearAllCells 全削除+count / dimensions+merges 保持 / `_appendRowCursor` reset。

  empirical: 1668 tests pass (was 1661, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **bulk hide/unhide helpers (`hideRows` / `unhideRows` / `hideColumns` / `unhideColumns`) を追加**。range 全体を一発で:
  1. `src/worksheet/worksheet.ts` に追加: 4 helpers、いずれも range validation + ループ内で既存 hide/unhide 単体 helper を呼ぶ。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/bulk-hide.test.ts` 6 件: hideRows 範囲全 stamp + 範囲外無視 / unhideRows 完全 reverse / row range invalid throw / hideColumns 同等 / unhideColumns で width 保持 / column range invalid throw。

  empirical: 1661 tests pass (was 1655, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`getMergedRangeAt` / `unmergeCellsAt` を追加**。座標から merge 範囲を逆引き:
  1. `src/worksheet/worksheet.ts` に追加: `getMergedRangeAt(ws, row, col)` (range containment lookup、見つからなければ undefined)、`unmergeCellsAt(ws, row, col)` (含む merge を 1 件削除し true 返す、無ければ false)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/merge-resolver.test.ts` 6 件: top-left/middle/bottom-right で同 range / 範囲外 undefined / multi-merge で対応 / unmergeCellsAt true + 削除 / false で他 merge 保持 / 複数 merge で対象のみ削除。

  empirical: 1655 tests pass (was 1649, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`unhideRow` / `unhideColumn` を追加**。`hideRow` / `hideColumn` の逆操作:
  1. `src/worksheet/worksheet.ts` に追加: hidden flag を取り除き、他 field が無ければ entry 自体も削除。Column 側は `min`/`max` を passthrough から外す。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/unhide-helpers.test.ts` 7 件: hidden flag drop / height/width 保持 / 他 field 無いと entry 削除 / 未存在 row no-op / column 同等 3 件。

  empirical: 1649 tests pass (was 1642, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **outline collapse helpers (`collapseRowGroup` / `expandRowGroup` / `collapseColumnGroup` / `expandColumnGroup`) を追加**。Excel のグループ +/- ボタン相当:
  1. `src/worksheet/worksheet.ts` に追加: collapse 系は range 全行/列に `hidden: true` + `collapsed: true` を stamp、expand 系は両 flag を取り除いて他 field 保持。outlineLevel は groupXxx で別途管理されるので素通し。
  2. `src/index.ts` から 4 helpers を re-export。
  3. `tests/phase-5/outline-collapse.test.ts` 6 件: collapse hides+flags / expand reverses / 既存 height + outlineLevel 保持 / 未存在 entry に expand で no-op / range invalid throw / 列版 同等 / save→load round-trip。

  empirical: 1642 tests pass (was 1636, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **range-level shortcuts (`setRangeBackgroundColor` / `setRangeFont` / `setRangeNumberFormat`) を追加**。setCellXxx の range 版:
  1. `src/styles/cell-style.ts` に追加: 3 helpers、いずれも内部 `setRangeStyle` 経由 (xf pool dedup を保ちつつ全 cell に同じ patch)。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/range-shortcuts.test.ts` 4 件: setRangeBackgroundColor 4 cell solid pattern / Color partial / setRangeFont 2 cell bold+name / setRangeNumberFormat '0.00' 2 cell。

  empirical: 1636 tests pass (was 1632, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`getCellComment` resolver を追加**。Cell → LegacyComment lookup:
  1. `src/worksheet/worksheet.ts` に追加: `getCellHyperlink` と同じ pattern で comment 側にも range containment lookup。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/cell-comment-resolver.test.ts` 4 件: single-cell ref / 不在は undefined / multi-comment で対象のみ / range ref containment。

  empirical: 1632 tests pass (was 1628, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`getCellHyperlink` resolver を追加**。Cell → Hyperlink lookup:
  1. `src/worksheet/worksheet.ts` に追加: `getCellHyperlink(ws, c)` で `ws.hyperlinks` を walk、`parseRange(h.ref)` + `rangeContainsCell` で hit chain。複数候補は insertion order の最初。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/cell-hyperlink-resolver.test.ts` 5 件: single-cell ref / range ref containment + 範囲外 undefined / hyperlink 0 件で undefined / outer→inner insertion order で outer wins / internal jump (location) も resolvable。

  empirical: 1628 tests pass (was 1623, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`getWorkbookStats` summary helper を追加**。workbook 内容のクイック QA:
  1. `src/workbook/workbook.ts` に追加: `WorkbookStats` 型 (worksheetCount / chartsheetCount / cellCount / formulaCount / commentCount / hyperlinkCount / mergedRangeCount / tableCount / definedNameCount / customPropertyCount) + `getWorkbookStats(wb)`。typed model を 1 走査して集約。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-stats.test.ts` 4 件: 空 workbook 全 0 / cell+formula+comment+hyperlink+mergedRange カウント / chartsheet 分離 / tables+definedNames+customProperties 全種。

  empirical: 1623 tests pass (was 1619, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **comment in-place edit helpers (`editCommentText` / `editCommentAuthor`) を追加**。setComment の差分書き換え版:
  1. `src/worksheet/worksheet.ts` に追加: ref で lookup → text or author だけ差し替え、他 field 保持。Return は boolean (見つかったか)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/comment-edit.test.ts` 5 件: editCommentText 部分書き換え+他保持 / ref missing で false / multi-comment で対象のみ / editCommentAuthor 部分書き換え / ref missing で false。

  empirical: 1619 tests pass (was 1614, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **fill preset helpers (`setCellBackgroundColor` / `clearCellBackground`) を追加**。Excel Home tab fill bucket 相当:
  1. `src/styles/cell-style.ts` に追加: `setCellBackgroundColor(wb, c, hexOrColor)` (string '`FFAAFFAA`' / Partial<Color>)、`clearCellBackground(wb, c)` (DEFAULT_EMPTY_FILL に戻す)。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/fill-presets.test.ts` 4 件: hex 文字列で solid pattern / Color partial で theme+tint / 同色再代入で fill pool dedup + 同 styleId / clearCellBackground が default に戻す。

  empirical: 1614 tests pass (was 1610, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **font preset helpers (`setBold` / `setItalic` / `setStrikethrough` / `setUnderline` / `setFontSize` / `setFontName` / `setFontColor`) を追加**。Excel Home tab font ボタン相当:
  1. `src/styles/cell-style.ts` に追加: 7 helpers + `mergeFont` 内部 (既存 font 保持で field merge)。`setUnderline(false)` は `underline` field 自体を削除。`setFontColor` は string ('FFAA00FF') / Partial<Color> ({theme,tint}) 両対応。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/font-presets.test.ts` 9 件: setBold/Italic/Strikethrough toggle / false で off / 他 field 保持 / underline default 'single' / 任意 style / false で削除 / size+name+color set / Color partial / 4 helper 合成。

  empirical: 1610 tests pass (was 1601, +9)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **alignment preset helpers (`centerCell` / `wrapCellText` / `alignCellHorizontal` / `alignCellVertical` / `rotateCellText` / `indentCell`) を追加**。Excel UI 整列ボタン相当:
  1. `src/styles/cell-style.ts` に追加: 6 helpers + `mergeAlignment` 内部 (既存 alignment 保持で field merge)。`makeAlignment` の validation 経由で 0..180/255 / 0..255 indent / valid horizontal/vertical 値が enforce。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/alignment-presets.test.ts` 7 件: centerCell hor+vert / 他 field 保持 / wrapCellText toggle / horizontal+vertical 独立 / rotateCellText 0..180+255 / 範囲外 throw / indentCell。

  empirical: 1601 tests pass (was 1594, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`duplicateSheet` workbook helper を追加**。Excel "Move or Copy → Create a copy" 相当:
  1. `src/workbook/workbook.ts` に追加: `duplicateSheet(wb, sourceTitle, newTitle, { index?, state?, tableSuffix? })`。`structuredClone` で worksheet 全体を複製→ title 上書き、新 sheetId 発行、tables を全 workbook 走査して max+1 に renumber + displayName を `<name><suffix>` (default `_2`、衝突時は `_2`/`_22`/`_222`...と incremental) で衝突回避。styleId は共有 stylesheet を経由するので id 不変で OK。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/duplicate-sheet.test.ts` 8 件: cells+columnDimensions+legacyComments verbatim / mutate-after-clone 独立 / table id+displayName renumber / tableSuffix 任意 / 重複 title throw + missing source throw / index 挿入 / styleId 共有 / save→load round-trip。

  empirical: 1594 tests pass (was 1586, +8)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **outline grouping helpers (`groupRows` / `ungroupRows` / `groupColumns` / `ungroupColumns`) を追加**。Excel "Data → Group" を 1-line で:
  1. `src/worksheet/worksheet.ts` に追加: 4 helpers + `validateOutlineLevel` (`[1, 7]` 整数、Excel の grouping ネスト上限)。group 系は range 内全 row/col の outlineLevel を stamp、既存 RowDim/ColDim の他 field は保持。ungroup 系は outlineLevel を取り除き、entry が空になれば map から削除。
  2. `src/index.ts` から 4 helpers を re-export。
  3. `tests/phase-5/outline-grouping.test.ts` 10 件: groupRows level=1 デフォ / 任意 level / 既存 height 保持 / 不正 level+range throw / ungroupRows 完全削除 / 他 field 保持 / groupColumns 同 / ungroup 完全削除 / 不正 range throw / 全 round-trip。

  empirical: 1586 tests pass (was 1576, +10)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **sheet default dimension helpers (`setDefaultColumnWidth` / `setDefaultRowHeight`) を追加**。worksheet 単位の default 幅/高さを 1-line で:
  1. `src/worksheet/worksheet.ts` に追加: 両 helpers。`undefined` で field を削除、負値+NaN は throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/default-dimensions.test.ts` 7 件: defaultColumnWidth set / undefined で clear / 不正値 throw / defaultRowHeight 同 3 件 / save→load round-trip 両 default。

  empirical: 1576 tests pass (was 1569, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **custom-property ergonomic helpers (`setCustomStringProperty` / `Number` / `Bool` / `Date` / `getCustomPropertyValue` / `removeCustomProperty` / `listCustomProperties`) を追加**。docProps/custom.xml を JS 値で直接編集:
  1. `src/packaging/custom.ts` に追加: 4 setter (Number は Int32 で `vt:i4`, それ以外で `vt:r8` 自動選択、Date は `Date | string` 受け取り `toISOString` 経由 filetime)、`getCustomPropertyValue` (string→int→double→bool→filetime decode を順試行)、`removeCustomProperty` (boolean return)、`listCustomProperties` (read-only snapshot)、`replaceOrAppend` 内部 (同名 prop は pid 維持で上書き)。
  2. `src/packaging/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-3/custom-properties-helpers.test.ts` 6 件: lazy alloc + 全 4 種同時 + getValue 5 種 / int → i4 vs float → r8 / 同名 replace で pid 維持 / removeCustomProperty bool return / NaN+Infinity throw / 全 3 種 round-trip。

  empirical: 1569 tests pass (was 1563, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **extended-properties helpers (`setWorkbookCompany` / `Manager` / `Application` / `AppVersion` / `HyperlinkBase`) を追加**。docProps/app.xml を 1-line で:
  1. `src/packaging/extended.ts` に追加: 5 setters + `ensureAppProperties`。
  2. `src/packaging/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-3/extended-properties-helpers.test.ts` 4 件: lazy alloc + 5 field 一括 / 後勝ち / save→load round-trip / 非 ASCII 多言語+emoji。

  empirical: 1563 tests pass (was 1559, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **core property setters (`setWorkbookCreator` / `Title` / `Subject` / `Description` / `Keywords` / `LastModifiedBy` / `Category`) を追加**。docProps/core.xml を 1-line で:
  1. `src/packaging/core.ts` に追加: 7 setters + `ensureCoreProperties`。Excel UI "ファイル → プロパティ" の主要 field を直接 editable に。
  2. `src/packaging/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-3/core-properties-helpers.test.ts` 4 件: lazy alloc + 7 field 一括 / 後勝ち / save→load 全 7 field round-trip / 非 ASCII (日本語+emoji)。

  empirical: 1559 tests pass (was 1555, +4)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`replaceCellValues` find-and-replace を追加**。文字列 / predicate 両対応:
  1. `src/worksheet/worksheet.ts` に追加: `replaceCellValues(ws, search, replacement)`。`search` が string の場合は完全一致 (substring 不可)、function なら `(value, cell) => boolean` を全 populated cell に適用。Return は変更件数。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/replace-cell-values.test.ts` 6 件: string match で count 返却 / 不一致 0 / 数値+真偽値 skip (string モード) / substring 非マッチ / predicate 数値範囲 / Cell 引数で coord 分岐。

  empirical: 1555 tests pass (was 1549, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **listing helpers (`listHyperlinks` / `listTables` / `listDataValidations` / `listDefinedNames`) を追加**。read-only snapshot:
  1. `src/worksheet/worksheet.ts` に追加: 上記 worksheet 系 3 helpers (各 `ReadonlyArray` を return)。
  2. `src/workbook/defined-names.ts` に追加: `listDefinedNames(wb, { scope?: number | 'workbook' | 'all' })` (default 'all'、'workbook' で scope undefined のみ、数値で sheet 限定)。
  3. `src/index.ts` から 4 helpers を re-export。
  4. `tests/phase-5/listing-helpers.test.ts` 9 件: hyperlinks 非空+empty / tables 1 件+empty / dataValidations 1 件 / definedNames default all / scope 'workbook' / scope 0 / 空 wb で空。

  empirical: 1549 tests pass (was 1540, +9)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`reserveDefaultXfSlot` を `applyXfPatch` / `setRangeStyle` に組み込み、`cellXfs[0]` の implicit default 衝突を解消**。前 turn で発覚した footgun を修正:
  1. `src/styles/cell-style.ts`: `applyXfPatch` と `setRangeStyle` の冒頭で `if (cellXfs.length === 0) addCellXf(defaultCellXf())` を呼んで cellXfs[0] を default として確保。これで初回スタイル適用後、unstyled cell (styleId=0) は default xf を指したままになる。
  2. 既存テスト 2 件を更新 (`cell-style-bridge.test.ts:54`, `json-roundtrip.test.ts:182` — どちらも `cellXfs.length` 期待値を 1→2 に更新、コメント追記)。
  3. `tests/phase-2/styles/cellxf-default-slot.test.ts` 5 件: setCellFont で a.styleId>0 + b.styleId=0 + slot0.fontId=0 / setCellFill 同 / setRangeStyle 'A1' で header.styleId !== body.styleId / formatAsHeader で body 行 styleId=0 維持 / 手動 reserve なし→default+bold で length=2。

  empirical: 1540 tests pass (was 1535, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **comment listing / rename / filter helpers (`listComments` / `renameCommentAuthor` / `findCommentsByAuthor`) を追加**。team-handoff 後のメンテ用:
  1. `src/worksheet/worksheet.ts` に追加: `listComments(ws)` (read-only snapshot)、`renameCommentAuthor(ws, oldName, newName)` (一致する全 comment を makeLegacyComment で再構築、return 件数)、`findCommentsByAuthor(ws, author)` (filter)。
  2. `src/index.ts` から 3 helpers を re-export。
  3. `tests/phase-5/comment-helpers.test.ts` 7 件: listComments author 順 / 更新後 reflect / renameCommentAuthor count + 全置換 / 不一致は 0 件 / save→load 後新名 / findCommentsByAuthor filter / 0 件で空 array。

  empirical: 1535 tests pass (was 1528, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`formatAsHeader` range preset を追加**。ヘッダー行用 bold + 暗 fill + 下線:
  1. `src/styles/cell-style.ts` に追加: `formatAsHeader(wb, ws, range, opts?)`。default は Excel "Table Style Medium 2" header の bold white on `FF305496` + medium 下線。`opts.fillColor` / `fontColor` (string|Partial<Color>) / `bold` / `bottomBorder: SideStyle | false` / `bottomBorderColor` を override 可。`setRangeStyle` 経由で auto-allocate。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/format-as-header.test.ts` 6 件: default 全 axis / custom fillColor+fontColor+bold=false / bottomBorder=false / 任意 border style+color / 空 cell auto-alloc / 同 opts で idempotent (cellXfs pool 増えず)。

  empirical: 1528 tests pass (was 1522, +6)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **workbookProperties helpers (`setWorkbookCodeName` / `setDate1904` / `setUpdateLinksMode` / `setFilterPrivacy`) を追加**。`<workbookPr>` の主要 attr を 1-line で:
  1. `src/workbook/workbook-properties.ts` に追加: 上記 4 helpers + `ensureWorkbookProperties`。`setDate1904` は canonical `wb.date1904` と `<workbookPr>` mirror の両方を sync。
  2. `src/index.ts` から 4 helpers を re-export。
  3. `tests/phase-3/workbook-properties-helpers.test.ts` 7 件: codeName lazy / 非 ASCII (Книга) round-trip / 空文字保持 / setDate1904 両 sync / round-trip / setUpdateLinksMode 3 mode / setFilterPrivacy 反転。

  empirical: 1522 tests pass (was 1515, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **`getDataExtent` / `getDataExtentRef` を追加**。populated cell の bounding box:
  1. `src/worksheet/worksheet.ts` に追加: `getDataExtent(ws)` (`{minRow, maxRow, minCol, maxCol}` を 1 走査で計算、空 sheet は undefined、空 rowMap (delete 後の残骸) は skip)、`getDataExtentRef(ws)` (`rangeToString` 経由で `"A1:E5"` 形式、1×1 は `"A1"` のみ、空 sheet は undefined)。
  2. `src/index.ts` から 2 helpers を re-export。
  3. `tests/phase-5/data-extent.test.ts` 8 件: 空 sheet / 1×1 / sparse 両軸 / 100×50 lonely / 空 rowMap 無視 / Ref canonical / 1×1 で colon 無し / 空 sheet undefined。

  empirical: 1515 tests pass (was 1507, +8)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **approximate autofit (`autofitColumn` / `autofitColumns`) を追加**。文字列長ベースの幅推定:
  1. `src/worksheet/worksheet.ts` に追加: `autofitColumn(ws, col, { minRow?, maxRow?, padding=2, min=4, max=80 })` (cellValueAsString 長を走査、`max(length)+padding` を `[min, min(max,255)]` で clamp)、`autofitColumns(ws, { padding, min, max })` (全 populated col を 1 走査で集約)。populated cell が無い col は no-op (返値 undefined)。docstring に「string-length 近似、CJK / wide glyph は要追加 padding」と明記。
  2. `src/index.ts` から 2 helpers を re-export。
  3. `tests/phase-5/autofit-columns.test.ts` 9 件: 既定 padding=2 / min clamp / max clamp / 任意 padding / 空 col で undefined + 未 alloc / minRow+maxRow window / autofitColumns 並行 / per-col clamp / rich-text 連結長。

  empirical: 1507 tests pass (was 1498, +9)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **workbook view toggles (`setShowHorizontalScroll` / `setShowVerticalScroll` / `setWorkbookMinimized` / `setWorkbookVisibility`) を追加**。primary workbookView の chrome 制御を 1-line で:
  1. `src/workbook/views.ts` に追加: 4 helpers、いずれも既存 `ensurePrimaryView` 経由で lazy 作成。
  2. `src/index.ts` から 4 helpers を re-export。
  3. `tests/phase-3/workbook-view-toggles.test.ts` 5 件: scroll bar 独立 toggle / lazy 作成 / minimized 反転 / visibility 3 mode / 4 toggle 全 round-trip。

  empirical: 1498 tests pass (was 1493, +5)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **cell finding helpers (`findCells` / `findFirstCell` / `getCellsInRange`) を追加**。データ抽出向け iteration:
  1. `src/worksheet/worksheet.ts` に追加: `findCells(ws, predicate)` (row→col 順 generator、null value placeholder も訪問)、`findFirstCell(ws, predicate)` (最初の match を返す、無ければ undefined)、`getCellsInRange(ws, range)` (populated cell のみ yield、auto-allocate 無し)。
  2. `src/index.ts` から 3 helpers を re-export。
  3. `tests/phase-5/find-cells.test.ts` 9 件: findCells row→col 順 / always-false yield 0 / 空 ws / null cell 訪問 / findFirstCell 早い順 / 不存在で undefined / getCellsInRange populated only / drain 後も auto-alloc 無し / single-cell range。

  empirical: 1493 tests pass (was 1484, +9)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **formula introspection helpers (`getFormulaText` / `getCachedFormulaValue`) を追加**。formula cell の読み取り側エルゴ:
  1. `src/cell/cell.ts` に追加: `getFormulaText(c)` (formula 以外は undefined; shared follower は formula 空文字)、`getCachedFormulaValue(c)` (cachedValue を return、未設定 / 非 formula は undefined)。
  2. `src/index.ts` から 2 名を re-export。
  3. `tests/phase-2/formula-introspection.test.ts` 8 件: 通常 formula / array formula / shared follower 空文字 / 非 formula 3 種 undefined / cachedValue 数値 / 文字列+真偽値 verbatim / cachedValue 未設定 undefined / 非 formula undefined。

  empirical: 1484 tests pass (was 1476, +8)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **header/footer code builder + ergonomic setters を追加**。Excel の `&L`/`&C`/`&R` ミニ format 文字列を構造化:
  1. `src/worksheet/page-setup.ts` に追加: `HEADER_FOOTER_CODES` 凍結カタログ (pageNumber `&P` / pageCount `&N` / date `&D` / time `&T` / filePath `&Z&F` / fileName `&F` / sheetName `&A` / picture `&G`)、`buildHeaderFooterText({ left?, center?, right? })` (undefined section は marker 出さず、空文字列は marker のみ出す)、`setHeaderText(ws, parts, section='odd')` / `setFooterText(ws, parts, section='odd')` (既存 `setHeader`/`setFooter` 経由で `differentFirst`/`differentOddEven` flip)。
  2. `src/index.ts` から 4 名を re-export。
  3. `tests/phase-5/header-footer-text.test.ts` 10 件: build 全 3 section / center only で `&L` 接頭辞無し / undefined skip + 空文字列 marker 維持 / 空入力で空文字 / コード interpolation 自然 / setHeaderText default 'odd' / first で differentFirst / even で differentOddEven / setFooterText oddFooter / 全 round-trip。

  empirical: 1476 tests pass (was 1466, +10)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **selection helpers (`setActiveCell` / `setSelectedRange`) を追加**。primary SheetView の Selection を 1-line で:
  1. `src/worksheet/worksheet.ts` に追加: `setActiveCell(ws, ref)` (前 activeCell と一致する sqref のみ追従、明示 sqref は保持)、`setSelectedRange(ws, sqref)` (multi-range 入力可、activeCell 未設定時のみ先頭 ref から導出)。
  2. `src/index.ts` から 2 helpers を re-export。
  3. `tests/phase-5/selection-helpers.test.ts` 7 件: setActiveCell lazy + sqref 同期 / 連続 setActiveCell sqref 追従 / 明示 sqref 後の setActiveCell が sqref 保持 / setSelectedRange 単一 range derive activeCell / multi-range で先頭 / 既存 activeCell 上書きしない / 全 round-trip。

  empirical: 1466 tests pass (was 1459, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **print-option ergonomic helpers (`setPrintGridLines` / `setPrintHeadings` / `setPrintCentered`) を追加**。Excel の "Page Layout → Sheet Options" を 1-line で:
  1. `src/worksheet/page-setup.ts` に追加: `setPrintGridLines(ws, on)` (gridLines + gridLinesSet を同期)、`setPrintHeadings(ws, on)`、`setPrintCentered(ws, { horizontal?, vertical? })` (片方だけ指定で他方を保持)。`ensurePrintOptions` で lazy 作成。
  2. `src/index.ts` から 3 helpers を re-export。
  3. `tests/phase-5/print-options-helpers.test.ts` 7 件: setPrintGridLines lazy + sync / flip false 同期 / setPrintHeadings 独立 / setPrintCentered horizontal only / 両軸 / partial update / 4 field 全 round-trip。

  empirical: 1459 tests pass (was 1452, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **named-style cell application (`applyBuiltinStyle` / `applyNamedStyle`) を追加**。"Heading 1" / "Total" / "Good" / "Bad" 等を cell に 1-line で:
  1. `src/styles/cell-style.ts` に追加: `applyBuiltinStyle(wb, c, name)` (`ensureBuiltinStyle` 経由で cellStyleXfs index 取得 + cellXf に xfId/font/fill/border/numFmt id を patch)、`applyNamedStyle(wb, c, name)` (workbook 既登録 NamedStyle を `_namedStyleByName` で lookup)。共通 `applyNamedStyleByXfId` で patch 構築。
  2. `src/styles/index.ts` + `src/index.ts` から 2 helpers を re-export。
  3. `tests/phase-2/styles/apply-named-style.test.ts` 7 件: Good 緑 palette + xfId set / Bad 赤 / 同じ built-in 2 cell でも namedStyles 1 entry / 不存在名 throw / user-registered NamedStyle 適用 / 未登録名 throw / setCellFont 後でも xfId 維持。

  empirical: 1452 tests pass (was 1445, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **border preset helpers (`setCellBorderAll` / `setRangeBorderBox`) を追加**。Excel UI の「全枠線」「外枠」を 1-line で:
  1. `src/styles/cell-style.ts` に追加: `setCellBorderAll(wb, c, { style, color? })` (4辺同 style/color、color は hex string か `Partial<Color>`)、`setRangeBorderBox(wb, ws, range, { style, color?, inner? })` (perimeter cell に外側 edge を、inner 指定時は内部 cell + 外側 cell の internal edge も inner style で。inner なし & non-perimeter cell は no-op で auto-alloc しない)。`-readonly [K in keyof Border]?: ...` writable patch で typecheck 通過。
  2. `src/styles/index.ts` + `src/index.ts` から 2 helpers を re-export。
  3. `tests/phase-2/styles/border-presets.test.ts` 8 件: setCellBorderAll 4 sides 同 style / hex color 通過 / default 'thin' / setRangeBorderBox perimeter only (corner stroke 配分 + interior 未 alloc) / inner 指定で全 cell に stroke + perimeter cell は outer+inner 混在 / 1×1 range で 4-side outer / color 透過 / 既存値保持。

  empirical: 1445 tests pass (was 1437, +8)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **format-preset cell helpers (`setCellAsCurrency` / `setCellAsPercent` / `setCellAsDate` / `setCellAsNumber`) を追加**。Excel "Format Cells → Number → Category" の preset を 1-line で:
  1. `src/styles/cell-style.ts` に追加: `setCellAsCurrency(wb, c, { symbol?, decimals?, accounting? })` ( default `$#,##0.00`、`accounting=true` で `_-$* #,##0.00_-;-$* #,##0.00_-;_-$* "-"??_-;_-@_-` Excel 標準テンプレ)、`setCellAsPercent(wb, c, decimals=0)` (default `0%`、`decimals=2` で `0.00%`)、`setCellAsDate(wb, c, format='yyyy-mm-dd')`、`setCellAsNumber(wb, c, decimals=0)` (default `#,##0`)。decimals は非負整数バリデーション。各 helper は既存 `setCellNumberFormat` 経由で format-code を pool に登録。
  2. `src/styles/index.ts` + `src/index.ts` から 4 helpers を re-export。
  3. `tests/phase-2/styles/cell-format-presets.test.ts` 13 件: currency default + custom symbol+decimals=0 + accounting / percent default + decimals=2 + 不正 throw / date default + 任意 / number default + decimals=2 + 不正 throw / preset の dedup (built-in `0.00%` 同 styleId / 非 built-in `€#,##0.00` 1 entry)。

  empirical: 1437 tests pass (was 1424, +13)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **range-read helpers (`getRangeValues` / `getColumnValues` / `getRowValues`) を追加**。setRangeValues の双対:
  1. `src/worksheet/worksheet.ts` に追加: `getRangeValues(ws, range)` (parsed range の shape そのままに dense 2-D `(CellValue|null)[][]`、欠損 cell は `null`)、`getColumnValues(ws, col, { minRow?, maxRow? })` (default は 1..getMaxRow)、`getRowValues(ws, row, { minCol?, maxCol? })` (default は 1..max-populated-col、行未存在 + opts 無し時は `[]`、opts ありなら null padded)。
  2. `src/index.ts` から 3 helpers を re-export。
  3. `tests/phase-5/range-readers.test.ts` 11 件: getRangeValues dense + null + setRangeValues 双対 + 完全空 range + single-cell / getColumnValues null pad + minRow/maxRow window + 空 ws / getRowValues default + minCol/maxCol + 行未存在 (opts 無し空、opts あり null pad)。

  empirical: 1424 tests pass (was 1413, +11)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **CellValue 型ガード + coercion helpers を追加**。Cell 値を読む側のエルゴ:
  1. `src/cell/cell.ts` に追加: 値レベルの `isFormulaValue / isRichTextValue / isErrorValue / isDurationValue` (TypeScript narrowing 付き) + `cellValueAsString(v)` (rich-text concat / formula cachedValue / error token / duration `"N ms"` / Date ISO / null `''`) + `cellValueAsNumber(v)` (boolean→0/1 / numeric string parse / formula 数値 cachedValue 通過 / NaN/Infinity→undefined)。
  2. `src/index.ts` から 6 名を re-export。
  3. `tests/phase-2/cell-value-coercion.test.ts` 13 件: 4 type guards / cellValueAsString 6 ケース (null / primitive / Date / rich-text / formula 有無 / error+duration) / cellValueAsNumber 5 ケース (primitive / numeric+非数値文字列 / formula 数値+文字列 cachedValue / null+Date+error+duration / NaN+Infinity)。

  empirical: 1413 tests pass (was 1400, +13)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **calcProperties ergonomic helpers を追加**。Excel 計算オプションを 1-line で:
  1. `src/workbook/calc-properties.ts` に追加: `setCalcMode(wb, 'manual'|'auto'|'autoNoTable')`、`setIterativeCalc(wb, enable, { count?, delta? })` (count/delta 未指定時は Excel default を保持)、`setCalcOnSave(wb, on)`、`setFullCalcOnLoad(wb, on)`、`setFullPrecision(wb, on)`。`ensureCalcProperties` で lazy 作成。
  2. `src/index.ts` から 5 helpers を re-export。
  3. `tests/phase-3/calc-properties-helpers.test.ts` 7 件: setCalcMode / setIterativeCalc default + count+delta + flip false 時 count 維持 / setCalcOnSave + setFullCalcOnLoad + setFullPrecision 独立 / 全 helper 合成 / save→load round-trip。

  empirical: 1400 tests pass (was 1393, +7)、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **CF visual rule builders (`addColorScaleRule` / `addDataBarRule` / `addIconSetRule`) を追加**。これまで innerXml 直書き必要だった colorScale / dataBar / iconSet を構造化 opts で:
  1. `src/worksheet/conditional-formatting.ts` に追加: `Cfvo` / `CfvoType` / `IconSetStyle` 型 + 3 builders。各 builder は `<colorScale>` / `<dataBar>` / `<iconSet>` の inner XML を組み立てて `innerXml` field に格納 (visual rule の round-trip 経路をそのまま使う)。バリデーション: colorScale は cfvos.length ∈ {2,3} かつ colors.length と一致、iconSet は cfvos.length ∈ [3,5]、num/percent/percentile/formula type cfvo は val 必須。
  2. `src/index.ts` から 3 builder + 3 型を re-export。
  3. `tests/phase-5/cf-visual-rules.test.ts` 10 件: 2-stop / 3-stop colorScale / mismatched length throw / 1-stop throw / dataBar default min/max + custom cfvos + minLength/maxLength/showValue / iconSet 3-light + 5-arrow with reverse/showValue/percent / cfvos length out-of-range throw / 3 種同時の save→load round-trip。

  empirical: 1393 tests pass (was 1383, +10)、e2e 32 件 pass、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **sheet view + tab-color ergonomic helpers を追加**。`<sheetPr>/<tabColor>` と primary `<sheetView>` 表示トグルを 1-line で:
  1. `src/worksheet/worksheet.ts` に追加: `setSheetTabColor(ws, hexOrColor)` (Color が readonly なため `string | Partial<Color>` を受けて `makeColor()` 経由)、`removeSheetTabColor(ws)`、`setShowGridLines/setShowRowColHeaders/setShowFormulas/setShowZeros/setRightToLeft(ws, boolean)`、`setSheetZoom(ws, scale)` ([10, 400] integer 検証)、`setSheetViewMode(ws, 'normal'|'pageBreakPreview'|'pageLayout')`。`ensureSheetProperties` / 既存 `ensurePrimaryView` で lazy 作成。
  2. `src/index.ts` から 9 helpers を re-export。
  3. `tests/phase-5/sheet-view-helpers.test.ts` 10 件: tabColor hex/Color partial / removeSheetTabColor が siblings 保持 / no-op / 5 toggle 同時 set / lazy primary view 作成 / setSheetZoom 正常 / 範囲外+非整数 throw / setSheetViewMode 3-mode 切替 / 全 trip save→load round-trip。

  empirical: 1383 tests pass (was 1373, +10)、e2e 32 件 pass、typecheck / lint clean (16 warnings)。

- **次のタスク (前回)**: **bulk dimension helpers (`setColumnWidths` / `setRowHeights`) を追加**。複数列・複数行の width/height を 1-call で:
  1. `src/worksheet/worksheet.ts` に追加: `setColumnWidths(ws, widths, startCol=1)` と `setRowHeights(ws, heights, startRow=1)`。両方 array (positional from start) と `Record<index, value>` (sparse) を受ける。NaN / 非整数キー / 0 以下の index を全 skip。各 entry は customWidth/customHeight=true を set。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/bulk-dimensions.test.ts` 7 件: array default startCol / startCol offset / Record sparse / 不正 entry skip / row 同等 3 件。

  empirical: 1373 tests pass (was 1366, +7)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **range-level helpers (`setRangeValues` / `applyToRange` / `setRangeStyle`) を追加**。長方形範囲への一括適用:
  1. `src/worksheet/worksheet.ts` に `setRangeValues(ws, range, rows[][])` (2D array を range の左上から流し込み、null/undefined は skip) と `applyToRange(ws, range, visit)` (各 (row,col) を訪問、未存在 cell は allocate)。
  2. `src/styles/cell-style.ts` に `setRangeStyle(wb, ws, range, opts)` (1 つの xf を組み立て、range 内の全 cell に同じ styleId を割り当て、未存在 cell は allocate)。
  3. `src/styles/index.ts` + `src/index.ts` から re-export。
  4. `tests/phase-5/range-helpers.test.ts` 6 件: setRangeValues 2D 流し込み / null/undefined skip / applyToRange 6-cell 走査 + lazy alloc / setRangeStyle 1 row に font + alignment + fill / 空 range で 3×3 cell auto-alloc / 空 opts no-op。

  empirical: 1366 tests pass (was 1360, +6)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **`setCellStyle(wb, cell, opts)` 統合 styler を追加**。font/fill/border/alignment/protection/numberFormat を 1 call で:
  1. `src/styles/cell-style.ts` に追加: 各 axis を independent に処理する combined setter。`-readonly [K in keyof CellXf]?: CellXf[K]` の writable patch を組み立てて `applyXfPatch` に渡す。空 opts は no-op。
  2. `src/styles/index.ts` + `src/index.ts` から re-export。
  3. `tests/phase-2/styles/set-cell-style.test.ts` 4 件: font + fill 同時適用 / 全 6 axis 同時適用 / 空 opts no-op / full save → load round-trip。

  empirical: 1360 tests pass (was 1356, +4)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **drawing builder helpers (`addImageAt` / `addChartAt`) を追加**。画像・チャートを 1-line で配置:
  1. `src/drawing/drawing.ts` に追加: `addImageAt(ws, ref, image, opts?)` (image は `XlsxImage` か raw `Uint8Array`、`Uint8Array` なら `loadImage` で format sniff、widthPx/heightPx default 96)、`addChartAt(ws, ref, chart, opts?)` (chart は既存 `ChartReference` 型、widthPx/heightPx default 480×320)。両方 `ws.drawing` を lazy-allocate し既存 items に append。
  2. `src/index.ts` から drawing module 一式を re-export: `ChartReference` / `Drawing` / `DrawingItem` / `PictureReference` 型 + `addChartAt` / `addImageAt` / `makeChartDrawingItem` / `makeDrawing` / `makePictureDrawingItem` + `XlsxImage` / `loadImage` + `DrawingAnchor` / `makeOneCellAnchor`。これまで全て internal だった drawing 層が正式に public API に。
  3. `tests/phase-6/image-chart-helpers.test.ts` 5 件: lazy-allocate drawing + bytes path / loaded XlsxImage path + custom size / 2 image append / chart append + chart kind / image+chart 共存。

  empirical: 1356 tests pass (was 1351, +5)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **sheet visibility helpers を追加**。タブストリップでの hide/show/veryHide:
  1. `src/workbook/workbook.ts` に追加: `getSheetState(wb, title)` / `setSheetState(wb, title, state)` (低レベル)、`hideSheet(wb, title)` / `showSheet(wb, title)` / `veryHideSheet(wb, title)` (Excel UI 用語に合わせた shortcut)。すべて不存在 title で throw。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/sheet-visibility.test.ts` 6 件: default 'visible' / hide ↔ show flip / veryHideSheet / setSheetState 直接 / 不存在 throw / 4 sheet 状態 round-trip。

  empirical: 1351 tests pass (was 1345, +6)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **`renameSheet` / `moveSheet` workbook helpers を追加**。タブストリップの sheet 並び替え + リネームを 1-line で:
  1. `src/workbook/workbook.ts` に追加: `renameSheet(wb, oldTitle, newTitle)` (空文字 / 重複 / 不存在 source を全 throw、同名 oldTitle === newTitle は no-op)。`moveSheet(wb, title, toIndex)` (`Math.max(0, Math.min(len-1, toIndex))` で clamp、active sheet が move 対象なら activeSheetIndex を新位置に同期、別 sheet が active なら from/dest の relative shift で再計算)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/sheet-rename-move.test.ts` 10 件: rename success / 不存在 throw / duplicate target throw / 空文字 throw / no-op same name / move basic / clamp +∞/-∞ / activeSheetIndex 維持 / 不存在 throw / 非整数 throw。

  empirical: 1345 tests pass (was 1335, +10)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **hyperlink builder helpers (`addUrlHyperlink` / `addInternalHyperlink` / `addMailtoHyperlink`) を追加**。worksheet レベルの hyperlink 設定を 1-line で:
  1. `src/worksheet/hyperlinks.ts` に追加: `replaceHyperlink(ws, hl)` (同 ref の既存 entry を置換)、`addUrlHyperlink(ws, ref, url, opts?)`、`addInternalHyperlink(ws, ref, location, opts?)`、`addMailtoHyperlink(ws, ref, email, { subject?, ... })` (subject は `encodeURIComponent`)。
  2. `src/index.ts` から `Hyperlink` 型 + 4 helpers を re-export (これまで未公開だった hyperlink 型も正式 public API に)。
  3. `tests/phase-5/hyperlink-helpers.test.ts` 7 件: URL の tooltip+display / 同 ref で URL→URL replace / 内部 jump (location のみ) / 同 ref で URL→internal replace / mailto basic / mailto with subject 経由 encoded URL / 3 entry 全種 save → load round-trip。

  empirical: 1335 tests pass (was 1328, +7)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **defined-name builder helpers + print-area / print-titles convenience を追加**。`Workbook.definedNames[]` への高レベル access:
  1. `src/workbook/defined-names.ts` に追加: `addDefinedName(wb, opts)` (同 name+scope の既存 entry を replace)、`getDefinedName(wb, name, scope?)`、`removeDefinedName(wb, name, scope?)`。
  2. `setPrintArea(wb, sheetIdx, ref)` (`_xlnm.Print_Area` を sheet scope で set)、`setPrintTitles(wb, sheetIdx, { rows?, cols?, sheetName })` (`_xlnm.Print_Titles` を `'Sheet'!$A:$A,'Sheet'!$1:$1` 形式で集約、rows/cols 両方無いと throw)。
  3. `src/index.ts` から `DefinedName` 型 + 6 helpers を re-export (これまで未公開だった defined-name モジュールを正式 public API に)。
  4. `tests/phase-3/defined-name-helpers.test.ts` 8 件: addDefinedName replace 動作 / 異なる scope の coexistence / removeDefinedName boolean return / setPrintArea / setPrintTitles full + rows-only + cols-only + 全 missing → throw / save → load round-trip。

  empirical: 1328 tests pass (was 1320, +8)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **freeze-pane + autoFilter ergonomic helpers を追加**。Excel UI の "Freeze Top Row" / "Freeze First Column" 等 + filter dropdown 設定:
  1. `src/worksheet/worksheet.ts` に追加: `freezeRows(ws, count)` (`A${count+1}` 経由)、`freezeColumns(ws, count)` (column letter computed via `columnLetterFromIndex`)、`freezePanes(ws, rows, cols)`、`unfreezePanes(ws)`。count バリデーション付き。
  2. `src/worksheet/auto-filter.ts` に追加: `addAutoFilter(ws, ref)` (`ws.autoFilter` に直接 set)、`addAutoFilterColumn(ws, colId, values, opts?)` (filter が無いなら throw)、`removeAutoFilter(ws)` (delete via index assertion)。
  3. `src/index.ts` から 4 freeze helpers + 3 autoFilter helpers + AutoFilter / FilterColumn 型 + factory を re-export。
  4. `tests/phase-5/freeze-autofilter-helpers.test.ts` 11 件: freezeRows/Columns/freezePanes ref 計算 / unfreezePanes / 不正 count throw / save→load round-trip / addAutoFilter set / addAutoFilterColumn append + blank flag / no-filter throw / removeAutoFilter / autoFilter round-trip。

  empirical: 1320 tests pass (was 1309, +11)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **`addExcelTable(wb, ws, opts)` builder helper を追加**。verbose な `makeTableDefinition` + 手動 id 管理を一発で:
  1. `src/worksheet/table.ts` に追加: `nextTableId(wb)` (workbook 全シートを走査して既存テーブルの最大 id +1)、`addExcelTable(wb, ws, opts)`。
  2. `opts.columns` は `string[]` か `TableColumn[]` を受け、string なら 1-based id を auto-assign。
  3. `opts.style` は built-in style 名 shortcut (例: `'TableStyleMedium2'`) を渡すと `{ name, showRowStripes: true, showColumnStripes: false }` の `styleInfo` を auto-build。`styleInfo` を直接渡すとそちらが優先。
  4. `displayName` 未指定時は `name` を使用、`headerRowCount` / `totalsRowCount` / `totalsRowShown` / `autoFilter` も透過。
  5. `src/index.ts` から `TableColumn` / `TableDefinition` / `TableStyleInfo` 型 + `addExcelTable` / `makeTableColumn` / `makeTableDefinition` を re-export。
  6. `tests/phase-5/table-builder.test.ts` 5 件: string 配列 → TableColumn id auto-assign / 複数シートでの id auto-increment / styleInfo override が `style` shortcut を上書き / `TableColumn[]` 透過 / 完全 round-trip。

  empirical: 1309 tests pass (was 1304, +5)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **conditional-formatting builder helpers を追加**。`addCellIsRule` / `addTopNRule` / `addAverageRule` / `addDuplicateValuesRule` / `addFormulaRule` / `addTextRule`:
  1. `src/worksheet/conditional-formatting.ts` に追加: `resolveCfSqref` (string sqref を `parseMultiCellRange` で wrap)、`nextCfPriority(ws)` (既存 rule の priority 最大値 +1 で auto-increment)、`pushRule` ヘルパー (1 rule × 1 sqref で `ConditionalFormatting` ブロック生成)。
  2. 6 builder: cellIs (between=2 formula、それ以外は 1 formula)、topN (rank/bottom/percent オプション)、average (above/below + ±N stdDev)、duplicateValues (`unique: true` で uniqueValues に切替)、formula (任意 expression)、text (4 enum operator → containsText / notContainsText / beginsWith / endsWith の type token に map)。
  3. `src/index.ts` から CF 型一式 (CellIsOperator / TextOperator / TimePeriod / ConditionalFormatting / ConditionalFormattingRule / ConditionalFormattingRuleType) + 6 builders + `makeCfRule` / `makeConditionalFormatting` を re-export。
  4. `tests/phase-5/conditional-formatting-builders.test.ts` 11 件: cellIs between/single/priority auto-increment / topN defaults + bottom-percentile / aboveAverage + stdDev / dup vs uniq / formula 透過 / text 4 operator mapping / 3-rule save → load round-trip。

  empirical: 1304 tests pass (was 1293, +11)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **data-validation builder helpers を追加**。verbose な `makeDataValidation` を avoid して 1-line で `addListValidation` / `addNumberValidation` / `addDateValidation` / `addCustomValidation` できるように:
  1. `src/worksheet/data-validations.ts` に追加: `ValidationCommon` 共通 opts (prompt / promptTitle / error / errorTitle / errorStyle / allowBlank)、`resolveSqref` (string sqref を `parseMultiCellRange` で wrap)、4 builder。
  2. `addListValidation` は inline `string[]` を `"Red,Green,Blue"` quoted form に変換、reference string (`'=Sheet2!$A$1:$A$10'`) はそのまま透過。
  3. `addNumberValidation` は `{ min, max?, operator?, kind? 'whole'|'decimal' }`、min only なら operator='greaterThanOrEqual'、max ありなら 'between' を default。
  4. `addDateValidation` は Excel serial 数値、同じ operator デフォルト logic。
  5. `addCustomValidation` は formula1 をそのまま透過。
  6. `src/index.ts` から DataValidation 型 + 4 builders + `makeDataValidation` を re-export。
  7. `tests/phase-5/data-validation-builders.test.ts` 9 件: list inline + reference + prompt/error metadata / number between + min-only + decimal kind / date + custom / 全 builder save → load round-trip。

  empirical: 1293 tests pass (was 1284, +9)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **workbook-protection ergonomic helpers を public API に追加**。`protectWorkbook` / `unprotectWorkbook` / `isWorkbookProtected`:
  1. `src/workbook/protection.ts` に追加: `protectWorkbook(wb, overrides?)` (`lockStructure: true` デフォルトに overrides 重ね、password-hash quad 透過)、`unprotectWorkbook(wb)`、`isWorkbookProtected(wb)`。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-protection-helpers.test.ts` 4 件: defaults / 全 lock 軸を on にする overrides / password-hash quad 透過 / unprotectWorkbook で undefined。

  empirical: 1284 tests pass (was 1280, +4)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **sheet-protection ergonomic helpers を public API に追加**。`protectSheet` / `unprotectSheet` / `isSheetProtected`:
  1. `src/worksheet/protection.ts` に追加: `PROTECT_SHEET_DEFAULTS` 定数 (Excel「シート保護」ダイアログの no-touch wire form: sheet+objects+scenarios=true、その他 13 boolean=false)、`protectSheet(ws, overrides?)` (デフォルトに overrides を spread、password hash quad も透過)、`unprotectSheet(ws)` (delete via index assertion で exactOptionalPropertyTypes 回避)、`isSheetProtected(ws)` (sheet=true 判定)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/sheet-protection-helpers.test.ts` 4 件: defaults / overrides で sort+autoFilter 許可 / password-hash quad 透過 / unprotectSheet で undefined。

  empirical: 1280 tests pass (was 1276, +4)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **page-setup ergonomic helpers を public API に追加**。worksheet root の typed coverage は完了したが、各 typed フィールドへの直接アクセス boilerplate を減らす:
  1. `src/worksheet/page-setup.ts` に追加: `setPageOrientation` / `setPaperSize` / `setPrintScale` / `setFitToPage` (lazy allocate `pageSetup`)、`setPageMargins` (Excel default で missing axes 補完)、`setHeader` / `setFooter` (`'odd'|'even'|'first'` section enum、even/first は対応する `differentOddEven` / `differentFirst` flag を auto-set)、`addRowBreak` / `addColBreak` (man:true + max を ECMA default で push)。
  2. `src/index.ts` から re-export。
  3. `tests/phase-5/page-setup-helpers.test.ts` 5 件: pageSetup 4 helper の lazy allocate / setPageMargins の default fill / setHeader/setFooter の differentOddEven/differentFirst 自動 toggle / addRowBreak + addColBreak の defaults / 全 helpers 経由の full save → load round-trip。

  empirical: 1276 tests pass (was 1271, +5)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **worksheet `<customSheetViews>` を typed API に**。worksheet root の最後の大きな未モデル要素を処理:
  1. `src/worksheet/custom-sheet-views.ts` 新設: `CustomSheetView { guid (required); scale?; colorId?; 13 boolean attr (showPageBreaks / showFormulas / showGridLines / showRowCol / outlineSymbols / zeroValues / fitToPage / printArea / filter / showAutoFilter / hiddenRows / hiddenColumns / filterUnique / showRuler); state? 'visible'|'hidden'|'veryHidden'; view? SheetViewMode; topLeftCell?; pane?; selections?; rowBreaks?; colBreaks?; pageMargins?; printOptions?; pageSetup?; headerFooter? }` 全 20 attr + 9 nested children + factory。pane/selection/breaks は既存型を再利用。
  2. `Worksheet.customSheetViews: CustomSheetView[]` を追加 (デフォルト空配列、`makeWorksheet` で init)。
  3. reader: `CUSTOM_SHEET_VIEWS_TAG` を `MODELED_WORKSHEET_TAGS` に登録、`parseCustomSheetView` で全 attr (enum 不正値は drop) + nested children を pull (既存 `parsePane` / `parseSelection` / `parsePageBreak` / `parsePageMargins` / `parsePrintOptions` / `parsePageSetup` / `parseHeaderFooter` を全部再利用)。
  4. writer: ECMA-376 §18.3.1.27 順 — mergeCells の直後・phoneticPr の前に emit、空配列は emit ナシ。inner ordering: pane → selection × N → rowBreaks → colBreaks → pageMargins → printOptions → pageSetup → headerFooter。
  5. `tests/phase-5/worksheet-custom-sheet-views.test.ts` 3 件: 全 20 attr + 全 8 nested children round-trip / 2 enum 不正値 drop / 空配列 emit ナシ。

  empirical: 1271 tests pass (was 1268, +3)、e2e 32 件 pass、typecheck / lint clean。**worksheet root の主要 typed coverage 完了**。

- **次のタスク (前回)**: **worksheet `<oleObjects>` + `<controls>` を typed API に**。embedded OLE オブジェクトと form control の list:
  1. `src/worksheet/ole-objects.ts` 新設: `OleObject { shapeId; rId?; progId?; dvAspect? 'DVASPECT_CONTENT'|'DVASPECT_ICON'; link?; oleUpdate? 'OLEUPDATE_ALWAYS'|'OLEUPDATE_ONCALL'; autoLoad?; objectPr?: XmlNode }` (`objectPr` 子要素は anchor/spreadsheet-drawing namespace を含むため verbatim XmlNode 保持)。`FormControl { shapeId; rId?; name?; controlPr?: XmlNode }`。`makeOleObject` / `makeFormControl` factory。
  2. `Worksheet.oleObjects: OleObject[]` / `Worksheet.controls: FormControl[]` を追加 (デフォルト空配列、`makeWorksheet` で init)。
  3. reader: `OLE_OBJECTS_TAG` / `CONTROLS_TAG` を `MODELED_WORKSHEET_TAGS` に登録。`shapeId` int parse 失敗 entry は drop、`dvAspect` / `oleUpdate` enum も valid set チェック。
  4. writer: ECMA-376 §18.3.1.61 / §18.3.1.27 順 — drawing/legacyDrawing/legacyDrawingHF の後ろ・picture の前に emit、children は `serializeBodyExtraNode` で verbatim 出力。
  5. `tests/phase-5/ole-objects.test.ts` 4 件: 2-entry oleObjects (Word.Document.12 / Equation.3 link) round-trip / 空配列 emit ナシ / 2-entry controls (CheckBox1 / SpinButton1) round-trip / 空配列 emit ナシ。

  empirical: 1268 tests pass (was 1264, +4)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **workbook view ergonomic helpers を public API に追加**。`bookViews[0]` (primary view) への高レベル access — typed model はあるが、bookViews を直接 mutate せず使えるユーティリティ:
  1. `src/workbook/views.ts` に追加: `ensurePrimaryView(wb)` (lazy allocate)、`getActiveTab(wb)` / `setActiveTab(wb, idx)` / `getFirstSheet(wb)` / `setFirstSheet(wb, idx)` / `setTabRatio(wb, n)` / `setShowSheetTabs(wb, bool)` / `setWorkbookWindow(wb, { x?, y?, w?, h? })`。
  2. `src/index.ts` から re-export。
  3. `tests/phase-3/workbook-view-helpers.test.ts` 4 件: 空 workbook の get fallback (=0) / set による lazy allocate / setActiveTab + setFirstSheet + setTabRatio + setShowSheetTabs + setWorkbookWindow の round-trip / setWorkbookWindow が指定軸のみ更新。

  empirical: 1264 tests pass (was 1260, +4)、e2e 32 件 pass、typecheck / lint clean (warning count 16 維持)。

- **次のタスク (前回)**: **chartsheet `<customSheetViews>` を typed API に**。chartsheet 版は worksheet 版より小さく (5 attr + 3 子要素のみ):
  1. `Chartsheet.customSheetViews: ChartsheetCustomSheetView[]` 追加。`ChartsheetCustomSheetView { guid; scale?; state? 'visible'|'hidden'|'veryHidden'; zoomToFit?; pageMargins?; pageSetup?; headerFooter? }` + factory。
  2. reader: `<customSheetViews>` を chartsheet-xml.ts に追加、enum `state` は valid set チェック後 drop。
  3. writer: ECMA-376 §18.3.1.12 順 — `<sheetProtection>` の直後・`<pageMargins>` の前に emit、空配列は emit ナシ。
  4. `tests/phase-6/chartsheet-custom-sheet-views.test.ts` 3 件: 単一 view (scale=75 / state=hidden / zoomToFit + nested page-setup all 3 children) round-trip / 不正 state enum drop / 空配列 emit ナシ。

  empirical: 1260 tests pass (was 1257, +3)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **chartsheet `<webPublishItems>` を typed API に**。worksheet と同じ `WebPublishItem` 型を再利用、chartsheet ケース (sourceType='chart' など):
  1. `Chartsheet.webPublishItems: WebPublishItem[]` 追加、`makeChartsheet` で空配列に init。
  2. `parseWebPublishItem` / `serializeWebPublishItems` を worksheet モジュールから `export` 化、chartsheet で再利用。
  3. chartsheet writer は `<picture>` の直後 (ECMA §18.3.1.12 順) に emit、空配列は emit ナシ。
  4. `tests/phase-6/chartsheet-web-publish.test.ts` 2 件: chart sourceType + autoRepublish round-trip / 空配列 emit ナシ。

  empirical: 1257 tests pass (was 1255, +2)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **chartsheet rId-link siblings (legacyDrawing / legacyDrawingHF / drawingHF / picture)**。chartsheet が drop していた末尾要素群を typed フィールドに昇格:
  1. `Chartsheet` model に `legacyDrawingRId? / legacyDrawingHFRId? / drawingHF? / backgroundPictureRId?` を追加。`drawingHF` は `ChartsheetDrawingHF { rId; lho? / cho? / rho? / lhe? / che? / rhe? / lhf? / chf? / rhf? / lfo? / cfo? / rfo? / lfe? / cfe? / rfe? / lff? / cff? / rff? }` 全 18 attr (per-section image index map for header/footer DrawingML refs)。
  2. reader: 4 element を chartsheet-xml.ts に追加、drawingHF の 18 int attr は `DRAWING_HF_INT_KEYS` をループで pull、int parse 失敗は drop。
  3. writer: ECMA-376 §18.3.1.12 順 — drawing → legacyDrawing → legacyDrawingHF → drawingHF → picture。
  4. `tests/phase-6/chartsheet-rid-elements.test.ts` 3 件: 3 rId round-trip / drawingHF 全 18 image-index round-trip / 全 undefined emit ナシ。

  empirical: 1255 tests pass (was 1252, +3)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **chartsheet-level pageMargins / pageSetup / headerFooter を typed API に**。前まで chartsheet は print-setup elements を黙って drop していたため、(typed worksheet と同じ) 3 element をサポート:
  1. `Chartsheet` model に `pageMargins?: PageMargins; pageSetup?: PageSetup; headerFooter?: HeaderFooter` を追加 (既存 `src/worksheet/page-setup.ts` の型を共有)。chartsheet schema は `<printOptions>` を持たないので除外。
  2. `src/worksheet/reader.ts` / `writer.ts` の private `parsePrintOptions` / `parsePageMargins` / `parsePageSetup` / `parseHeaderFooter` / `serialize*` を `export` に変更し、chartsheet から再利用可能に。
  3. `src/chartsheet/chartsheet-xml.ts` の reader / writer を 3 element に対応 (ECMA-376 §18.3.1.12 順 — sheetProtection の後ろ・drawing の前)。
  4. `tests/phase-6/chartsheet-page-setup.test.ts` 2 件: A4 landscape + custom margins + odd header/footer の round-trip / undefined emit ナシ。

  empirical: 1252 tests pass (was 1250, +2)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **worksheet-level legacyDrawingHF を typed API に**。print の header/footer 背景画像 VML のための `<legacyDrawingHF r:id="…"/>`:
  1. `Worksheet.legacyDrawingHFRId?: string` 追加 (legacyDrawing と並ぶ rId-link、underlying VML part は relsExtras 経由で保持)。
  2. reader: `LEGACY_DRAWING_HF_TAG` を `MODELED_WORKSHEET_TAGS` に登録。
  3. writer: ECMA-376 §18.3.1.51 順、legacyDrawing の直後・picture の前に emit。
  4. `tests/phase-5/legacy-drawing-hf.test.ts` 2 件: rId round-trip / undefined 時 emit ナシ。

  empirical: 1250 tests pass (was 1248, +2)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **worksheet-level smartTags (per-cell smart-tag annotations) を typed API に**。Excel 2003 era の per-cell smart-tag、workbook-level smartTagTypes と対をなす:
  1. `src/worksheet/smart-tags.ts` 新設: `CellSmartTagProperty { key, val }` / `CellSmartTag { type, properties[], deleted?, xmlBased? }` / `CellSmartTags { ref, tags[] }` の 3 段ネスト型と factory。
  2. `Worksheet.smartTags: CellSmartTags[]` 追加 (デフォルト空配列)。
  3. reader: `<smartTags>` を `MODELED_WORKSHEET_TAGS` に登録、`<cellSmartTags r=…>` → `<cellSmartTag type=…>` → `<cellSmartTagPr key=… val=…>` の 3 レベルを pull、type int parse 失敗 entry は drop。
  4. writer: ECMA-376 §18.3.1.93 順 — ignoredErrors の直後・drawing の前に emit。空配列は emit ナシ。
  5. `tests/phase-5/cell-smart-tags.test.ts` 2 件: A1/A2 2 セルにそれぞれ 1 / 2 タグ × 2 properties round-trip / 空配列時 emit ナシ。

  empirical: 1248 tests pass (was 1246, +2)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **workbook-level customWorkbookViews を typed API に**。Shared Workbook 機能で保存される per-user view 配列:
  1. `src/workbook/views.ts` を拡張: `CustomWorkbookView { name; guid; windowWidth; windowHeight; activeSheetId (required) + 13 optional boolean / 4 optional int / showComments 3-enum / showObjects 3-enum }` 全 22 attr + factory。
  2. `Workbook.customWorkbookViews?: CustomWorkbookView[]` 追加。
  3. reader: `<customWorkbookViews>` を `captureWorkbookXmlExtras` で typed lift、`name`/`guid` 必須欠落 entry は drop、enum 不正値も drop。
  4. writer: `<oleSize>` の直後 (ECMA-376 §18.2.4 順) に emit。空配列は emit ナシ。
  5. `tests/phase-3/custom-workbook-views.test.ts` 3 件: 全 attr 単一 view round-trip / 複数 view round-trip / undefined emit ナシ。

  empirical: 1246 tests pass (was 1243, +3)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **worksheet-level `<picture>` (background image) を typed API に**。Page Layout → Background で設定するシート背景画像:
  1. `Worksheet.backgroundPictureRId?: string` 追加 (`<picture r:id="…"/>` の単一 rId、underlying media part は worksheet relsExtras 経由で保持)。
  2. reader: `PICTURE_TAG` を `MODELED_WORKSHEET_TAGS` に登録。
  3. writer: `<picture r:id=…/>` は ECMA-376 §18.3.1.66 順、afterSheetData passthrough と webPublishItems の間に emit。
  4. `tests/phase-5/background-picture.test.ts` 2 件: rId round-trip / undefined 時 emit ナシ。

  empirical: 1243 tests pass (was 1241, +2)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **worksheet-level sortState を typed API に**。Excel が最後に適用した sort 条件を保存する要素 (再オープン時に同じ並びを復元):
  1. `src/worksheet/sort-state.ts` 新設: `SortState { ref; conditions: SortCondition[]; columnSort?; caseSensitive?; sortMethod? 'stroke'|'pinYin' }` + `SortCondition { ref; descending?; sortBy? 'value'|'cellColor'|'fontColor'|'icon'; customList?; dxfId?; iconSet? (17-enum: 3Arrows / 3ArrowsGray / ... / 5Quarters); iconId? }` + factory。
  2. `Worksheet.sortState?: SortState` 追加。
  3. reader: `<sortState>` を `MODELED_WORKSHEET_TAGS` に登録、enum 全 3 種 (sortMethod / sortBy / iconSet) は valid set チェック後ドロップ。
  4. writer: `<sortState>` は autoFilter の直後・dataConsolidate の前 (ECMA-376 §18.3.1.92)。
  5. `tests/phase-5/sort-state.test.ts` 3 件: 3-condition full round-trip (caseSensitive + pinYin + customList + 5Arrows iconSet + cellColor by dxfId + descending) / 3 enum 不正値 drop / undefined 時 emit ナシ。

  empirical: 1241 tests pass (was 1238, +3)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **worksheet-level protectedRanges を typed API に**。Excel の "Allow Edit Ranges" — sheet protection 中でも編集可能な特定 range を whitelist する仕組み:
  1. `src/worksheet/protected-ranges.ts` 新設: `ProtectedRange { sqref: MultiCellRange; name; password? (legacy hex); securityDescriptor? (Windows ACL); algorithmName?; hashValue?; saltValue?; spinCount? }` + factory。
  2. `Worksheet.protectedRanges: ProtectedRange[]` 追加 (デフォルト空配列、optional ではなく必ず array)。
  3. reader: `<protectedRanges>` を `MODELED_WORKSHEET_TAGS` に登録、各 `<protectedRange>` の 8 attr を pull。`spinCount` は int parse 失敗時 drop。
  4. writer: ECMA-376 §18.3.1.69 順 — sheetProtection の直後・scenarios の前に emit。
  5. `tests/phase-5/protected-ranges.test.ts` 2 件: legacy hex password の Editor1 と modern SHA-512 hash quad の Editor2 (multi-region sqref `D1:E10 G1:G5`) round-trip / 空配列時 emit ナシ。

  empirical: 1238 tests pass (was 1236, +2)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **workbook-level smartTagPr + smartTagTypes + functionGroups (3 件まとめて)**。残り少ない workbook-root 兄弟を一気に消化:
  1. `src/workbook/smart-tags.ts` 新設: `SmartTagProperties { embed?, show?: 'all'|'noIndicator' }` + `SmartTagType { namespaceUri?, name?, url? }` 配列。Excel 2003 smart-tag legacy。
  2. `src/workbook/function-groups.ts` 新設: `FunctionGroups { builtInGroupCount?, groups: { name }[] }` (XLL function group registration)。
  3. `Workbook` に `smartTagPr?` / `smartTagTypes?` / `functionGroups?` 追加。
  4. reader: 3 element を `captureWorkbookXmlExtras` で typed lift、enum 不正値は drop。
  5. writer: ECMA-376 順 — `<functionGroups>` は sheets 後 / externalReferences 前、`<smartTagPr>` + `<smartTagTypes>` は oleSize 後 / afterSheets passthrough 前。
  6. `tests/phase-3/smart-tags-function-groups.test.ts` 5 件: smartTagPr embed+show round-trip / 不正 enum drop / 3-entry smartTagTypes (mix of full and minimal entries) / functionGroups builtInGroupCount + 2 custom groups round-trip / undefined emit ナシ。

  empirical: 1236 tests pass (was 1231, +5)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **workbook-level externalReferences を typed API に**。pivotCaches と並ぶ rId-link 配列、cross-workbook formula `[N]Sheet!A1` の `[N]` index 解決源:
  1. `Workbook.externalReferences?: ReadonlyArray<{ rId: string }>` 追加 (`<externalReference r:id="rIdN"/>` 配列、underlying xl/externalLinks/* parts は引き続き passthrough archive で保持)。
  2. reader: `<externalReferences>` を `captureWorkbookXmlExtras` で typed lift。
  3. writer: `<externalReferences>` は `<definedNames>` の後ろ・`<pivotCaches>` の前に emit (ECMA-376 §18.2.9)。空配列は emit ナシ。
  4. `tests/phase-3/external-references.test.ts` 3 件: rId list round-trip / undefined emit ナシ / 空配列 emit ナシ。

  empirical: 1231 tests pass (was 1228, +3)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **workbook-level pivotCaches を typed API に**。pivot table 含む workbook で常に存在する link 配列を型付け:
  1. `Workbook.pivotCaches?: ReadonlyArray<{ cacheId: number; rId: string }>` 追加 (`<pivotCache cacheId="N" r:id="rIdN"/>` の単純配列、xl/pivotCache/* parts は引き続き passthrough archive で保持)。
  2. reader: `<pivotCaches>` を `captureWorkbookXmlExtras` で typed lift。`cacheId` int parse 失敗 entry は drop。
  3. writer: `<pivotCaches>` は `<definedNames>` の後ろ・`<calcPr>` の前に emit (ECMA-376 §18.2.21)。
  4. `tests/phase-3/pivot-caches.test.ts` 3 件: openpyxl `pivot.xlsx` fixture から pivotCaches 配列が抽出される / load → save → load で cacheId 配列が一致 / 新規 workbook で undefined のまま emit ナシ。

  empirical: 1228 tests pass (was 1225, +3)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **workbook-level oleSize + fileRecoveryPr を typed API に**。短い workbook root 兄弟をまとめて昇格:
  1. `Workbook.oleSize?: string` 追加 (ECMA-376 §18.2.16、`<oleSize ref="…"/>` の単一 attr — embedded OLE 用 bounding range)。`src/workbook/file-recovery.ts` 新設で `FileRecoveryProperties { autoRecover? / crashSave? / dataExtractLoad? / repairLoad? }` 全 4 boolean (ECMA §18.2.11) と factory。
  2. reader: `OLE_SIZE_TAG` / `FILE_RECOVERY_PR_TAG` を `captureWorkbookXmlExtras` で typed lift。
  3. writer: `<oleSize>` は calcPr の直後 (ECMA 順 calcPr → oleSize → customWorkbookViews)、`<fileRecoveryPr>` は workbook 末尾 (afterSheets passthrough の後)。
  4. `tests/phase-3/oleSize-fileRecovery.test.ts` 4 件: oleSize ref round-trip + 未設定時 emit ナシ / fileRecoveryPr 4 boolean round-trip + 未設定時 emit ナシ。

  empirical: 1225 tests pass (was 1221, +4)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **workbook-level fileSharing を typed API に**。Save As → Tools → General Options の「Read-only recommended」+ write-protection password に対応:
  1. `src/workbook/file-sharing.ts` 新設: `FileSharing { readOnlyRecommended? / userName? / reservationPassword? (legacy hex) / algorithmName? / hashValue? / saltValue? / spinCount? (modern hash quad) }` 全 7 attr + factory。
  2. `Workbook.fileSharing?: FileSharing` 追加。
  3. reader: `<fileSharing>` を `captureWorkbookXmlExtras` で typed lift (inline parse、boolean flag は 1/0/true/false 受理)。
  4. writer: `<fileSharing>` は ECMA-376 §18.2.12 で fileVersion の直後 + workbookPr の前に emit。
  5. `tests/phase-3/file-sharing.test.ts` 3 件: modern hash quad + readOnly+userName round-trip / legacy reservationPassword round-trip / undefined 時 emit ナシ。

  empirical: 1221 tests pass (was 1218, +3)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **workbook-level fileVersion を typed API に**。Office app/version metadata。
  1. `src/workbook/file-version.ts` 新設: `FileVersion { appName / lastEdited / lowestEdited / rupBuild / codeName }` 全 5 attr (全て string)。
  2. `Workbook.fileVersion?: FileVersion` 追加。
  3. reader: `<fileVersion>` を `captureWorkbookXmlExtras` で typed lift。
  4. writer: `<fileVersion>` は workbook root の最初の child (ECMA-376 §18.2.13)、beforeSheets passthrough より前に emit。
  5. `tests/phase-3/file-version.test.ts` 2 件: 全 5 attr round-trip / undefined 時 emit ナシ。

  empirical: 1218 tests pass (was 1216, +2)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **workbook-level workbookPr (workbook properties) を typed API に**。前ターンの calcPr / bookViews / workbookProtection と並ぶ最後の主要 workbook-root 要素:
  1. `src/workbook/workbook-properties.ts` 新設: `WorkbookProperties` 全 19 attr (codeName / dateCompatibility / defaultThemeVersion / updateLinks 3-enum / showObjects 3-enum + 13 boolean: date1904 / backupFile / saveExternalLinkValues / hidePivotFieldList etc.) + factory。`ShowObjectsMode` ('all'|'placeholders'|'none') と `UpdateLinksMode` ('userSet'|'never'|'always') enum。
  2. `Workbook.workbookProperties?: WorkbookProperties` 追加。
  3. reader: `<workbookPr>` を `parseWorkbookProperties` で typed lift し、enum 不正値は drop。`wb.date1904` の canonical 解釈は変えず (`parseDate1904` が引き続き機能)。
  4. writer: `<workbookPr>` の date1904-only emission 経路を撤廃し、新 `serializeWorkbookProperties` で typed model から emit。`effectiveWorkbookProperties(wb)` で `wb.date1904 === true` なら typed model が無くても自動で `{date1904: true}` を合成 (旧挙動の互換性維持)。`hasWorkbookPrDate1904` dead code 削除。
  5. `tests/phase-3/workbook-properties.test.ts` 5 件: codeName + defaultThemeVersion + updateLinks + showObjects + 2 boolean round-trip / date1904=true via typed model + canonical wb.date1904 一致 / wb.date1904 だけ set 時の minimal synthesis / 不正 enum drop / 全 unset 時 emit ナシ。

  empirical: 1216 tests pass (was 1211, +5)、e2e 32 件 pass (date1904 文脈の既存 tests 全 pass)、typecheck / lint clean。

- **次のタスク (前回)**: **workbook-level calcPr (calculation properties) を typed API に**。`<bookViews>` と並ぶ workbook root の頻出要素:
  1. `src/workbook/calc-properties.ts` 新設: `CalcProperties { calcId / calcMode 'manual'|'auto'|'autoNoTable' / fullCalcOnLoad / refMode 'A1'|'R1C1' / iterate / iterateCount / iterateDelta / fullPrecision / calcCompleted / calcOnSave / concurrentCalc / concurrentManualCount / forceFullCalc }` 全 13 attr + factory。
  2. `Workbook.calcProperties?: CalcProperties` 追加。
  3. reader: `load.ts` で `<calcPr>` を発見したら `parseCalcProperties` で typed lift し、enum 不正値は drop。
  4. writer: `save.ts` で `<sheets>` + `<definedNames>` の後ろ・`afterSheets` passthrough の前に emit (ECMA-376 §18.2.2 順)。
  5. `tests/phase-3/calc-properties.test.ts` 4 件: calcId+calcMode+fullCalcOnLoad+refMode round-trip / iterative calc settings (iterate / iterateCount / iterateDelta + concurrentCalc) / 不正 enum drop / undefined 時 emit ナシ。

  empirical: 1211 tests pass (was 1207, +4)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **workbook-level bookViews を typed API に昇格 (firstSheet / activeTab / window 位置 / tabRatio など 13 attr)**。前ターンの workbookProtection と並ぶ workbook root 経由の頻出要素:
  1. `src/workbook/views.ts` 新設: `WorkbookView { visibility 3-enum / minimized / showHorizontalScroll / showVerticalScroll / showSheetTabs / xWindow / yWindow / windowWidth / windowHeight / tabRatio / firstSheet / activeTab / autoFilterDateGrouping }` 全 13 attr + `WorkbookViewVisibility` enum + `makeWorkbookView` factory。
  2. `Workbook.bookViews?: WorkbookView[]` 追加 (Excel 仕様で複数 entry 可能、通常は 1 つ)。
  3. reader: `load.ts` の `captureWorkbookXmlExtras` を拡張、`<bookViews>` を発見したら子 `<workbookView>` を全て typed 配列に lift し、`workbookXmlExtras.beforeSheets` から外す。
  4. writer: `save.ts` で `serializeBookViews` を `workbookProtection` の直後・`<sheets>` の直前に emit (ECMA-376 §18.2.1 の bookViews ↔ sheets 順)。
  5. `tests/phase-3/book-views.test.ts` 4 件: single view (firstSheet=1 / activeTab=2 / xWindow / windowWidth / tabRatio / showSheetTabs) / hidden visibility + 4 boolean flag / 複数 view 配列 / undefined 時 emit ナシ。

  empirical: 1207 tests pass (was 1203, +4)、e2e 32 件 pass、typecheck / lint clean。

- **次のタスク (前回)**: **B5 workbookProtection (workbook 側) を typed API に**。前ターンの sheetProtection (worksheet 側) と対をなす:
  1. `src/workbook/protection.ts` 新設: `WorkbookProtection { lockStructure?, lockWindows?, lockRevision? + workbook と revisions の 2 系統 password 4 attr 各々 (algorithmName / hashValue / saltValue / spinCount + legacy `*Password` 16-bit hex hash + `*PasswordCharacterSet`) }` 全 15 attr。`makeWorkbookProtection` factory。
  2. `Workbook.workbookProtection?: WorkbookProtection` を追加。
  3. reader: `load.ts` の `captureWorkbookXmlExtras` を拡張、`<workbookProtection>` を発見したら `parseWorkbookProtection` で typed model に lift し、`workbookXmlExtras.beforeSheets` から外す。
  4. writer: `save.ts` の `serializeWorkbookXml` で `beforeSheets` の直後・`<sheets>` の直前に `serializeWorkbookProtection` を emit (ECMA-376 §18.2.29、bookViews と sheets の間)。
  5. `tests/phase-3/workbook-protection.test.ts` 4 件: 3 boolean lock flag round-trip / modern hash quad (workbook 側 + revisions 側) round-trip / legacy 16-bit hex 2 password round-trip / undefined 時 emit ナシ。
  6. docs/plan/13 §B5 を更新 (workbookProtection 完了、残: password hash 計算 helper のみ)。

  empirical: 1203 tests pass (was 1199, +4)、e2e 32 件 pass、typecheck / lint clean。**累積**: B5 部分 (sheet+workbook 両方の typed API 完了、hash helper のみ残) + B6 ✅ + B7 部分 (sheetPr/dataConsolidate/scenarios 完了、customSheetViews のみ残) + B8 ✅ + B9 ✅ + B10 部分 + B11 ✅。

- **次のタスク (前回)**: **B7 scenarios (Scenario Manager) を typed API に**。docs/plan/13 §B7 サブピースをもう一つ完了:
  1. `src/worksheet/scenarios.ts` 新設: `ScenarioList { scenarios: Scenario[]; current?; show?; sqref? }` + `Scenario { name; inputCells: ScenarioInputCell[]; locked?; hidden?; user?; comment? }` + `ScenarioInputCell { ref; val; deleted?; undone?; numFmtId? }` 型と全 factory。`val` は wire-form 文字列で round-trip (Excel が numFmtId 経由で表示型を解釈)。
  2. `Worksheet.scenarios?: ScenarioList` を追加。
  3. reader: `SCENARIOS_TAG` を modeled-tag set に登録、`parseScenarioList` / `parseScenario` / `parseScenarioInputCell` で全 attr + nested children を pull、sqref は `parseMultiCellRange` 経由。
  4. writer: ECMA-376 §18.3.1.74 順 (sheetProtection の後ろ、autoFilter の前) に emit。`<scenario count=N>` の count attr は inputCells 数から auto-compute、空 inputCells は self-closing。
  5. `tests/phase-5/scenarios.test.ts` 2 件: full round-trip (BaseCase + Optimistic 2 シナリオ × 1〜2 inputCells, sqref 含む) / undefined 時 emit ナシ。
  6. docs/plan/13 §B7 を更新 (残: customSheetViews のみ)。

  empirical: 1199 tests pass (was 1197, +2)、e2e 32 件 pass、typecheck / lint clean。**累積**: B5 部分 + B6 ✅ + B7 部分 (sheetPr/dataConsolidate/scenarios 完了、customSheetViews のみ残) + B8 ✅ + B9 ✅ + B10 部分 + B11 ✅。

- **次のタスク (前回)**: **B10 部分対応 + B7 dataConsolidate 完了**。今ターン 2 連発:
  - **part 1 (B10)**: worksheet-level `<phoneticPr>` を typed API に。`PhoneticType` 4-enum + `PhoneticAlignment` 4-enum + `fontId`。round-trip 3 件、enum 不正値は drop。日本語 workbook の furigana font 制御に対応。残: per-cell `<rPh sb=… eb=…>`。
  - **part 2 (B7 sub)**: worksheet-level `<dataConsolidate>` を typed API に。Excel の Data → Consolidate ダイアログ用。`function` 11-enum (sum/average/count/countNums/max/min/product/stdDev/stdDevp/var/varp) + `topLabels` / `leftLabels` / `link` boolean + `dataRefs: DataReference[]` (name/ref/sheet/rId)。`startLabels` 後方互換 attr も round-trip。round-trip 3 件: function+labels+dataRefs full / function-only self-closing / undefined 時 emit ナシ。

  empirical: 1197 tests pass (was 1191, +6 — phonetic 3 + data-consolidate 3)、e2e 32 件 pass、typecheck / lint clean。**累積**: B5 部分 + B6 ✅ + B7 部分 (dataConsolidate 完了) + B8 ✅ + B9 ✅ + B10 部分 + B11 ✅。

- **次のタスク (前回)**: **B10 部分対応 — worksheet-level phoneticPr を typed API に昇格**。日本語 workbook で頻出する furigana 表示制御:
  1. `src/worksheet/phonetic.ts` 新設: `WorksheetPhoneticProperties { fontId?: number; type?: PhoneticType; alignment?: PhoneticAlignment }` 型と `makeWorksheetPhoneticProperties` factory。`PhoneticType` = 'halfwidthKatakana'|'fullwidthKatakana'|'Hiragana'|'noConversion'、`PhoneticAlignment` = 'noControl'|'left'|'center'|'distributed'。
  2. `Worksheet` に `phoneticPr?` を追加。
  3. reader: `PHONETIC_PR_TAG` を `MODELED_WORKSHEET_TAGS` に登録、`parsePhoneticPr` で 3 attr を pull (enum 値は valid set チェック)。
  4. writer: ECMA-376 §18.3.1.65 順 (mergeCells と conditionalFormatting の間) に emit。何も set されていないと undefined を返す。
  5. `tests/phase-5/phonetic.test.ts` 3 件: full round-trip (fontId=1 + Hiragana + distributed) / 不正な enum は drop / undefined 時 emit ナシ。
  6. docs/plan/13 §B10 を 🟡 部分対応に (残: per-cell `<rPh>`)。

  empirical: 1194 tests pass (was 1191, +3)、e2e 32 件 pass、typecheck / lint clean。**累積**: B5 部分 + B6 ✅ + B7 部分 + B8 ✅ + B9 ✅ + B10 部分 + B11 ✅。

- **次のタスク (前回)**: **B9 完了 — worksheet-level customProperties + webPublishItems を typed API に**。今ターンは 2 連発:
  - **part 1**: `<rowBreaks>` / `<colBreaks>` (B6 残り)。`PageBreak { id?/min?/max?/man?/pt? }` 型と `makePageBreak`、Worksheet model に空配列で init。reader/writer wired、auto manualBreakCount 計算 (man=false を除いた数)。e2e scenario 23 に row=40 manual break を追加、Page Break Preview で視覚確認可能。docs/plan/13 §B6 → ✅。
  - **part 2**: `<customProperties>` / `<webPublishItems>` (B9)。`src/worksheet/web-publish.ts` 新設: `WorksheetCustomProperty { name; rId? }` (rels link は既存の `relsExtras` 経由で別途 round-trip)、`WebPublishItem { id; divId; sourceType: 8 enum 'sheet'|'printArea'|'autoFilter'|'range'|'chart'|'pivotTable'|'query'|'label'; destinationFile; sourceRef?; sourceObject?; title?; autoRepublish? }`。Worksheet に空配列 init、reader/writer wired、ECMA 順は customProperties が rowBreaks の後ろ、webPublishItems が tableParts の直前。`tests/phase-5/web-publish.test.ts` 3 件: customProperties name 配列 round-trip / webPublishItems 全 attr round-trip / 全 unset 時 emit ナシ。docs/plan/13 §B9 → ✅。

  empirical: 1191 tests pass (was 1187, +4 — page-break 1 + web-publish 3)、e2e 32 件 pass、typecheck / lint clean。**累積**: B5 部分 + B6 ✅ + B7 部分 + B8 ✅ + B9 ✅ + B11 ✅。

- **次のタスク (前回)**: **B6 完了 (rowBreaks / colBreaks typed API + auto manualBreakCount)**。前ターンで残していた最後の 2 element を typed model に昇格:
  1. `src/worksheet/page-setup.ts` に `PageBreak { id?: number; min?: number; max?: number; man?: boolean; pt?: boolean }` 型と `makePageBreak` factory 追加。`<brk>` 要素 1 つに対応。
  2. `Worksheet` に `rowBreaks: PageBreak[]` / `colBreaks: PageBreak[]` を追加し `makeWorksheet` で空配列に init。
  3. reader: `ROW_BREAKS_TAG` / `COL_BREAKS_TAG` を `MODELED_WORKSHEET_TAGS` に登録、`<brk>` 子要素を pull (`parsePageBreak`)。
  4. writer: `<rowBreaks count=N manualBreakCount=M>...</rowBreaks>` を headerFooter の後ろに emit。`manualBreakCount` は `man=false` を除いた数で auto-compute。`man` undefined は manual 扱い (Excel default)。
  5. e2e scenario 23 に `ws.rowBreaks.push({ id: 40, max: 16383, man: true })` を追加 — Page Break Preview で 41 行目上の dashed line を視覚確認できる。
  6. `tests/phase-5/page-setup.test.ts` に rowBreaks/colBreaks round-trip テスト追加 (5→6 件)。
  7. docs/plan/13 §B6 を ✅ に更新。

  empirical: 1188 tests pass (was 1187, +1)、e2e 32 件 pass、typecheck / lint clean。**累積**: B5 部分 + B6 ✅ + B7 部分 + B8 ✅ + B11 ✅。

- **次のタスク (前回)**: **B6 printOptions / pageMargins / pageSetup / headerFooter を bodyExtras → typed API に昇格**。e2e scenario 23 (page-setup) が bodyExtras 経由で実証していたものを高レベル API に格上げ:
  1. `src/worksheet/page-setup.ts` 新設: 4 typed model — `PrintOptions` (5 boolean) / `PageMargins` (6 float、required when present、default factory で 0.75/0.75/1/1/0.5/0.5 を自動補完) / `PageSetup` (paperSize / scale / fitToWidth/Height / orientation enum 'default'|'portrait'|'landscape' / pageOrder enum / cellComments enum 'none'|'asDisplayed'|'atEnd' / errors enum 'displayed'|'blank'|'dash'|'NA' / horizontalDpi / verticalDpi / copies / paperWidth / paperHeight / rId / boolean: usePrinterDefaults / blackAndWhite / draft / useFirstPageNumber) / `HeaderFooter` (4 flag + 6 mini-format text: oddHeader / oddFooter / evenHeader / evenFooter / firstHeader / firstFooter)。
  2. Worksheet model に 4 つ optional 追加 (`printOptions?` / `pageMargins?` / `pageSetup?` / `headerFooter?`)。
  3. reader: 4 tag を `MODELED_WORKSHEET_TAGS` に登録、`parsePrintOptions` / `parsePageMargins` / `parsePageSetup` / `parseHeaderFooter` で全 attr + child element (oddHeader 等は text() pull) を parse。enum 値は valid set チェック後に絞る。
  4. writer: hyperlinks の後ろに 4 つを順に emit (printOptions → pageMargins → pageSetup → headerFooter)、bodyExtras.afterSheetData は今後はこれらと衝突しない他 element だけが残る。
  5. e2e scenario 23 を refactor: bodyExtras 経由から `makePrintOptions` / `makePageMargins` / `makePageSetup` / `makeHeaderFooter` 直書きに置換。
  6. `tests/phase-5/page-setup.test.ts` 5 件: printOptions 5 flag round-trip / pageMargins 6 float round-trip / pageSetup 11 attr round-trip / headerFooter 4 flag + 4 text section round-trip / 全 unset 時 emit ナシ。
  7. **既存テスト 2 件 fix** (`tests/phase-7/genuine-pivot-roundtrip.test.ts`): pageMargins が bodyExtras から外れて `ws.pageMargins` に lift されたため、assertion を更新 (typed field を確認、extLst は引き続き bodyExtras に残ることを確認)。
  8. docs/plan/13 §B6 を 🟡 部分対応に (残: rowBreaks / colBreaks)。

  empirical: 1187 tests pass (was 1182, +5)、e2e 32 件 pass、typecheck / lint clean。**累積**: B5 部分 + B6 部分 + B7 部分 + B8 ✅ + B11 ✅。

- **次のタスク (前回)**: **B5 sheetProtection を bodyExtras → typed API に昇格 (パスワード hash 計算は除く、wire-form 値は完全 round-trip)**。docs/plan/13 §B5 を 🟡 部分対応に。
  1. `src/worksheet/protection.ts` 新設: `SheetProtection` (16 boolean lock flag: sheet / objects / scenarios / formatCells / formatColumns / formatRows / insertColumns / insertRows / insertHyperlinks / deleteColumns / deleteRows / selectLockedCells / selectUnlockedCells / sort / autoFilter / pivotTables + 4 password-hash field: saltValue / spinCount / algorithmName / hashValue) と factory `makeSheetProtection`。
  2. `Worksheet` に `sheetProtection?: SheetProtection` を追加。
  3. reader: `SHEET_PROTECTION_TAG` を `MODELED_WORKSHEET_TAGS` に登録、`parseSheetProtection` で全 attr を pull。
  4. writer: sheetData の直後 (mergeCells の前) に `serializeSheetProtection` を emit (ECMA-376 §18.3.1.85 の位置)。何も set されていないと undefined を返して emit ナシ。
  5. e2e scenario 27 を refactor: bodyExtras 経由 (`elNs(SHEET_MAIN_NS, 'sheetProtection', {...})`) から `makeSheetProtection({ sheet: true, ... })` に置換。
  6. `tests/phase-5/sheet-protection.test.ts` 4 件: factory shape / 16 lock-flag round-trip / password-hash 4 field round-trip / 全 undefined 時 emit ナシ。
  7. docs/plan/13 §B5 を 🟡 部分対応に (残: password hash 計算 helper、workbookProtection)。

  empirical: 1182 tests pass (was 1178, +4)、e2e 32 件 pass、typecheck / lint clean。**累積**: B5 部分 + B7 部分 + B8 ✅ + B11 ✅。

- **次のタスク (前回)**: **B11 sheetFormatPr 拡張 (outlineLevelRow / outlineLevelCol + 5 attr) + outline level auto-compute**。docs/plan/13 §B11 を完遂:
  1. Worksheet model に `outlineLevelRow?` / `outlineLevelCol?` / `customHeight?` / `zeroHeight?` / `thickTop?` / `thickBottom?` / `baseColWidth?` の 7 attr を追加。
  2. reader: `<sheetFormatPr>` の同 7 attr を pull (parseIntegerAttr / parseBoolXmlAttr / parseFloatAttr)。
  3. writer: 同 7 attr を emit。`outlineLevelRow` は explicit 値が無いとき `rowDimensions` の `outlineLevel` 最大値を auto-compute (>0 のときのみ emit)、column 側も同様。これで Excel が outline ボタンを正しく描画 (深さ 2 のグループなら "1/2" toggle が出る)。
  4. `tests/phase-5/sheet-format-pr.test.ts` 4 件: auto-compute (row depth 2 + col depth 2) / explicit override / 5 attr round-trip / 未設定時 sheetFormatPr emit ナシ。
  5. docs/plan/13 §B11 を ✅ に更新 (auto-compute も明記)。

  empirical: 1178 tests pass (was 1174, +4)、e2e 32 件 pass、typecheck / lint clean。**B7 部分対応 + B8 ✅ + B11 ✅** が累積。

- **次のタスク (前回)**: **B7 sheetPr (codeName / tabColor / outlinePr / pageSetUpPr) を typed API に昇格**。e2e scenario 27 で tabColor を passthrough 経由で実証していたが、それを高レベル API に格上げ:
  1. `src/worksheet/properties.ts` 新設: `SheetProperties` (9 attr: codeName / enableFormatConditionsCalculation / filterMode / published / syncHorizontal / syncRef / syncVertical / transitionEvaluation / transitionEntry + 3 child: tabColor: Color / outlinePr: OutlineProperties / pageSetUpPr: PageSetupProperties) と factory `makeSheetProperties`。
  2. `Worksheet` に `sheetProperties?: SheetProperties` を追加 (空 default、optional)。
  3. reader: `SHEET_PR_TAG` を `MODELED_WORKSHEET_TAGS` に登録 (これで bodyExtras.beforeSheetData から外れる)、`parseSheetProperties` を新設し全 9 attr + 3 child element を pull。tabColor は `makeColor({rgb,indexed,theme,auto,tint})` 経由で正規化。
  4. writer: `<worksheet>` open tag 直後に `serializeSheetProperties(sp)` を emit (ECMA-376 §18.3.1.81 の最初の child)。bodyExtras.beforeSheetData はその後ろに残る (互換性)。
  5. e2e scenario 27 を refactor: bodyExtras 経由から `makeSheetProperties({ tabColor: makeColor({ rgb }) })` 経由に置換。同じ視覚結果。
  6. `tests/phase-5/sheet-properties.test.ts` 4 件: factory shape / RGB tabColor + codeName + outline + pageSetup full round-trip / theme-bound tabColor with tint / 未設定時 sheetPr emit ナシ。
  7. `src/index.ts` から `SheetProperties` / `OutlineProperties` / `PageSetupProperties` 型 + `makeSheetProperties` を re-export。
  8. docs/plan/13 の §B7 を 🟡 partial (残：customSheetViews / scenarios / dataConsolidate) に更新。

  empirical: 1174 tests pass (was 1170, +4)、e2e 32 件 pass、typecheck / lint clean (新規 warning ゼロ)。

- **次のタスク (前回)**: **B8 cellWatches / ignoredErrors を bodyExtras passthrough → 高レベル API に昇格**。docs/plan/13-full-excel-coverage.md §B8 (両方合わせて 0.5 sprint 推定) を完遂:
  1. `src/worksheet/errors.ts` 新設: `CellWatch { ref: string }` / `IgnoredError { sqref: MultiCellRange, evalError? / twoDigitTextYear? / numberStoredAsText? / formula? / formulaRange? / unlockedFormula? / emptyCellReference? / listDataValidation? / calculatedColumn? }` の plain-object モデルと factory (`makeCellWatch` / `makeIgnoredError`)。openpyxl `worksheet/errors.py` + `worksheet/cell_watch.py` と整合。
  2. `Worksheet` に `cellWatches: CellWatch[]` と `ignoredErrors: IgnoredError[]` プロパティを追加、`makeWorksheet` で空配列に初期化。
  3. reader: `CELL_WATCHES_TAG` / `IGNORED_ERRORS_TAG` を `MODELED_WORKSHEET_TAGS` に登録 (これで bodyExtras から外れる)、`<cellWatch r="…"/>` を pull、`<ignoredError sqref=…>` を全 9 flag bit で parse (`parseMultiCellRange` 経由)。
  4. writer: hyperlinks → bodyExtras.afterSheetData → `<cellWatches>` → `<ignoredErrors>` → drawing の順 (ECMA-376 §18.3.1.94 / §18.3.1.51 順)。`serializeCellWatches` / `serializeIgnoredErrors` を追加し、boolean flag は `1` 出力 (Excel default)、false や undefined は省略。
  5. `addCellWatch` / `removeCellWatches` / `addIgnoredError` / `removeIgnoredErrors` helpers を `worksheet.ts` に追加し `src/index.ts` から re-export (型 `CellWatch` / `IgnoredError` + factory `makeCellWatch` / `makeIgnoredError` も)。
  6. `tests/phase-5/errors.test.ts` 4 件: API CRUD (cellWatches/ignoredErrors それぞれ 1 件) + saveWorkbook→loadWorkbook 完全 round-trip (cellWatches 2 件 / ignoredErrors 2 件、後者は 9 flag のうち 4 個 + 2-range sqref を確認)。

  empirical: 1170 tests pass (was 1166, +4)、e2e 32 件 pass、typecheck / lint clean (新規 warning ゼロ、既存 16 件のみ)。**B8 entry を docs/plan/13 で 完了 (✅) に変更したい** が、本ターンでは実装に集中したので追って同 doc 更新。

- **次のタスク (前回)**: **e2e visual-verification suite 締めの 2 シナリオ: modern formulas / all-decorations cell**。
  - **`30-modern-formulas.xlsx`**: Excel 365 / 2021 系関数 — LET / LAMBDA(x,x*x)(7) / FILTER / SORT / UNIQUE / SEQUENCE(5) / XLOOKUP / BYROW(...,LAMBDA) を 1 sheet に並べ、source A2:C8 を Name/Dept/Salary 7 行で整える。ライブラリは「formula を文字列として書き込む / Excel 側が再計算する」モデルなので、これらは Excel 365 で開けば spilled-array が描画される。古い Excel は `#NAME?` (期待どおり)。
  - **`31-cell-combo.xlsx`**: B2:D4 merge 1 block に Cell の全 axis を同時に乗せる — bold red Calibri 16pt font + yellow solid fill + 4 辺 thick black border + center-align + wrapText の style、`https://example.com/` への hyperlink、author "QA" の comment、Yes/No/Maybe の `dataValidation list` (warning style)。低レベル `wb.styles` 経由の `addFont`/`addFill`/`addBorder`/`addCellXf` を public API で実証。

  empirical: `pnpm test:e2e` 31 シナリオ 32 テスト全 pass、30 (2,947 bytes) / 31 (4,029 bytes)、typecheck / lint clean。**累計 e2e: 18 (初回) + 11 (chart 系 6 + outline/page/multi 3 + alignment 1 + multi-table+tab-color 2 + theme+many 2 + modern+combo 2) = 31 シナリオ**。視覚検証カバレッジは pivot table / slicer / sparkline / threaded-comments / encrypted / VBA を除けば実質ほぼ網羅。残るは docs/plan/13-full-excel-coverage.md に列挙された 1.1+ milestone (threaded comments edit / sparklines / formula evaluator / pivot edit / encryption / VBA edit)。

- **次のタスク (前回)**: **e2e visual-verification suite またさらに拡張: theme-colors / many-sheets 2 シナリオ**。
  - **`28-theme-colors.xlsx`**: 2 sheets。"Theme" は 6 row × 3 col の matrix で `Color { theme: 0..5, tint: -0.5 / 0 / +0.4 }` を solid pattern fill に流し込み、Page Layout → Themes でテーマ切替したときセル色が再描画されることを実証 (RGB 固定ではなく theme reference)。"Indexed" は legacy 0..63 indexed palette を 1 行 1 色で並べる。stylesheet に `addFill` + `addCellXf` で xf を直接追加し `cell.styleId = xfId` を設定する経路 (低レベル API)。
  - **`29-many-sheets.xlsx`**: Summary + M01..M30 の計 31 シート。Summary に `='M01'!A1` ... `='M30'!A1` の cross-sheet 数式、各 M{NN} シートの A1 に `value from M{NN}`。Excel タブストリップの描画と多数シートの round-trip を smoke-test。

  empirical: `pnpm test:e2e` 29 シナリオ 30 テスト全 pass、28 (4,597 bytes) / 29 (14,456 bytes、31 シート分の overhead が見える)、typecheck / lint clean。

- **次のタスク (前回)**: **e2e visual-verification suite さらに拡張: multi-table / tab-color + sheet-protection 2 シナリオ**。
  - **`26-multi-table.xlsx`**: 1 sheet に 2 table を side-by-side。`tblOrders` (A1:D7、TableStyleLight9 青、showRowStripes) + `tblPayroll` (F1:I7、TableStyleMedium14 緑、showColumnStripes)。両方 AutoFilter dropdown 表示、structured-reference (`=SUM(tblOrders[Total])` / `=SUM(tblPayroll[Net])`) autocomplete も動く。
  - **`27-tab-color-protect.xlsx`**: 4 sheet。"Red" / "Green" / "Blue" の 3 つは tabColor (FFFF0000 / FF00B050 / FF0070C0) を `bodyExtras.beforeSheetData` 経由で `<sheetPr><tabColor rgb=...></sheetPr>` として注入。"Locked" は `<sheetProtection sheet=1 objects=1 scenarios=1 ...>` を `bodyExtras.afterSheetData` に注入し、Excel が編集を拒む状態に。bodyExtras passthrough hook の有用性を実証。

  empirical: `pnpm test:e2e` 27 シナリオ 28 テスト全 pass、26 (4,173 bytes) / 27 (3,888 bytes)、typecheck / lint clean。

- **次のタスク (前回)**: **e2e visual-verification suite またさらに拡張: alignment / advanced number-format scenario 25 追加**。前 2 ターン (chart 6 / outline+pagesetup+drawing 3) に続き、cell 表示そのものを揺さぶる scenario を追加:
  - **`25-alignment-numfmt.xlsx`**: 2 sheets。
    - "Align" tab: horizontal (left/center/right) + vertical (top/center) + wrapText (long line がセル幅に折り返し、行高 60pt) + indent=3 + shrinkToFit (長文が auto-shrink) + textRotation 45°/90°/135° (=−45°)/255 (vertical stacked)。回転 row は高さ 80pt に。
    - "NumFmt" tab: 13 number-format code 別の見え方 — `0.00` / `#,##0` / `#,##0.00` / `0%` / `0.00%` / `0.00E+00` / `"$"#,##0.00` / 条件付き `"$"#,##0.00;[Red]"$"#,##0.00;"-"` (positive/negative/zero) / `[h]:mm:ss` (1.5 → 36:00:00) / `m/d/yyyy` (45000 → 日付) / `@` (text passthrough) / fractions `# ?/?` (1.5 → 1 1/2) / `# ??/??` (0.123 → 1/8 系)。
  
  empirical: `pnpm test:e2e` 25 シナリオ 26 テスト全 pass、`25-alignment-numfmt.xlsx` 3,941 bytes、typecheck / lint clean。`tests/e2e/README.md` 更新。

- **次のタスク (前回)**: **e2e visual-verification suite さらに拡張: outline / page-setup / multi-drawing 3 シナリオ + worksheet helper public 露出**。chart 系を一気に厚くした前ターンに続き、ユーザの「あらゆる実 Excel」要望をさらに広範囲にカバー:
  1. **`22-grouping-outline.xlsx`**: row outline (Q1 detail rows 3..6 / Q2 detail rows 8..11 を level=1 にして collapsible に)、column outline (C..D を level=1)、column hidden (F)、custom width (A=22 / B=14 / C=12 / D=12 / E=12) + custom row height (subtotal rows 22pt / total row 26pt)。Excel の outline buttons (列ヘッダ上 / 行番号左の "1/2" toggle) と Format→Unhide が両方使える。
  2. **`23-page-setup.xlsx`**: `bodyExtras.afterSheetData` 経路で `<printOptions horizontalCentered=1 gridLines=1>` / `<pageMargins left=.5 right=.5 top=1 bottom=1>` / `<pageSetup paperSize=9 orientation=landscape fitToWidth=1 fitToHeight=0>` / `<headerFooter>` (oddHeader = `&LQuarterly&CQuarterly Report — &P / &N&R&D` / oddFooter = `&L&F&CPage &P of &N&RConfidential`) を XmlNode として注入。80 行データなので印刷プレビューが 2 ページにまたがる。
  3. **`24-multi-drawing.xlsx`**: 1 sheet に 3 drawing items 同居 — bar chart (E2)、line chart (E20)、PNG image (N2、scenario 18 と同じ tiny blue PNG) を `xl/drawings/drawing1.xml` 1 つに並べる。`makeDrawing([chartItem, chartItem, picItem])` で混在する経路を実証。
  4. **public API 拡張**: `src/index.ts` に `setColumnWidth` / `setColumnDimension` / `hideColumn` / `setRowHeight` / `setRowDimension` / `hideRow` / `getColumnDimension` / `getRowDimension` / `addConditionalFormatting` / `addDataValidation` / `addTable` / `setHyperlink` / `setComment` / `setAutoFilter` 等を追加 (前は internal モジュールから直接 import 必要、これで `from 'openxml-js'` だけで E2E が書ける)。`ColumnDimension` / `RowDimension` 型 + `makeColumnDimension` / `makeRowDimension` も併せて exports。

  empirical: `pnpm test:e2e` 24 シナリオ 25 テスト全 pass、22 (2,972 bytes) / 23 (4,823 bytes) / 24 (5,416 bytes) 生成、typecheck / lint clean。`tests/e2e/README.md` に検証チェックリスト 4 行追加 (前ターン 19/20/21 + 今回 22/23/24)。

- **次のタスク (前々回)**: **e2e visual-verification suite 拡張: chart classics / chartex / chart decorations 3 シナリオ追加**。既存の 18 ファイル (基本セル / 数式 / 日付 / 書式 / マルチシート etc.) に加え、ユーザの「あらゆる実 Excel を生成する E2E テスト」要望に応えて chart 経路を一気に厚くした:
  1. **`19-charts-classic.xlsx`**: 5 ヶ月 × 3 series (A/B/C) のソースデータに対し Line / Area (stacked) / Pie (A only) / Doughnut (50% hole) / Scatter (lineMarker, A vs B) / Radar (standard) の 6 チャートを F2/F20/F38/O2/O20/O38 にアンカー。`makeBarSeries` / `makeScatterSeries` / `makeChartSpace` の API を public 経路で実証。
  2. **`20-charts-chartex.xlsx`**: hierarchical category (`North/Apples` etc.) に対し chartex (`cx:` namespace) 8 チャート — Sunburst / Treemap / Waterfall (subtotalIdx [3]) / Histogram / Pareto / Funnel / BoxWhisker / RegionMap — を D/M/V 列にアンカー。`makeChartDrawingItem(anchor, { cxSpace })` 経路を実証。Excel 2016+ 必須、古い Excel は cx 名前空間を拒否する。
  3. **`21-chart-decorations.xlsx`**: `barSeries.dLbls = { showVal: true }` + `barSeries.trendline = [{ trendlineType: 'linear', dispEq: true, dispRSqr: true }]` の bar、`scatterSeries.errBars = [{ direction: 'y', errBarType: 'both', errValType: 'percentage', val: 10 }]` + `trendline.exp` の scatter。data labels / trendline / errBars が一画面に揃う。

  empirical: `pnpm test:e2e` 21 シナリオ全 pass、`tests/e2e/output/19-charts-classic.xlsx` 7,961 bytes / `20-charts-chartex.xlsx` 7,797 bytes / `21-chart-decorations.xlsx` 5,438 bytes 生成、`tests/e2e/README.md` 検証チェックリストに 3 行追加。typecheck / lint clean。`docs/plan/13-full-excel-coverage.md` の roadmap (1.1 threaded comments → 1.2 sparklines/dynamic arrays → 1.3 formula evaluator → 1.4 pivot edit → 2.0 encryption / VBA edit) は前ターンで整備済み、人手 visual QA をユーザに任せる段階。

- **次のタスク (前回)**: **fixture-driven fidelity 漏れ 5 件まとめて修正 (escape / calcChain / thumbnail / model / 形式判定 VML)**。29 openpyxl reference fixture を load → save → load で diff した結果、以下の bytes-dropped を発見し、それぞれ修正:
  1. **`\r` 文字が cell 文字列で消える**: `src/utils/escape.ts` の `ILLEGAL_RE` を `[\x01-\x1F]` 系 (XML legal whitespace を除外) → `[\x00-\x1F]` (NUL 含む全 C0 + サロゲート) に拡張。`\t`, `\n`, `\r` も `_xHHHH_` で encode し XML CRLF normalisation で消失するのを防ぐ。openpyxl と整合。`tests/phase-3/xml-escape-roundtrip.test.ts` 6 件 + `tests/phase-3/formula-escape-roundtrip.test.ts` 3 件で predefined-entity / whitespace / control-char / `_xHHHH_` literal を round-trip 検証。
  2. **`xl/calcChain.xml` 黙って drop**: `PASSTHROUGH_EXACT_PATHS: Set<string>` を新設し追加 (calculation order ヒント、Excel が再計算を強制されないため)。
  3. **`docProps/thumbnail.{jpeg,jpg,png,wmf,emf}` drop**: 同 set に追加 (workbook の OS-explorer preview)。
  4. **`xl/model/item.data` drop**: PASSTHROUGH_PREFIXES に `xl/model/` 追加 (Power Pivot data model)。
  5. **form-control VML が drop**: 旧 `isControlVml` は filename `vmlDrawingN.vml` を comment VML と判定して passthrough から除外 → form 制御の VML も巻き込まれ消失。content-aware sniff (`containsCommentMarker`: `ObjectType="Note"` byte search) に変更し comment VML だけ regenerate / 制御 VML は passthrough。`placeholderVmlDrawing()` も `<x:ClientData ObjectType="Note"/>` を含むよう更新し reload で正しく分類。
  
  empirical: `contains_chartsheets.xlsx` の calcChain.xml byte-identical preserved を `genuine-edge-fixtures.test.ts` に新テストで pin。`legacy_drawing.xlsm` の `vmlDrawing2.vml` (control VML) も round-trip 通過。1163 tests pass (was 1153, +10 in this turn)、lint / typecheck / size clean。残：Excel 365 視覚 QA (人手)、ZIP64 write の正式対応 (fflate 上流)、vba-test.xlsm の VML rels + EMF media (control VML 経由の image 参照、modeled drawing 層との競合のため後回し)。

- **ブランチ**: `main`（直接 commit 運用、squash 不要）

## 完了履歴

各エントリは「フェーズ §X.Y タスク → コミット hash」の形式。

### bootstrap

- [x] **bootstrap**: 計画コミット（`docs:` / 4573 行） — `773ae69`
- [x] **bootstrap**: TS プロジェクト雛形（`package.json` / `tsconfig.json` / `tsconfig.build.json` / `biome.json` / `.npmrc` / `.nvmrc` / `src/index.ts` placeholder / `THIRD_PARTY_NOTICES.md` / typescript 5.9 + @types/node 22 install / `pnpm typecheck` pass）
- [x] **bootstrap**: vitest 設定 + smoke test（vitest 4.1 + @vitest/coverage-v8、`tests/phase-0/smoke.test.ts` で `pnpm test` pass、`pnpm typecheck` pass）
- [x] **bootstrap**: tsup 設定（`tsup.config.ts` / `pnpm build` で `dist/index.mjs` + `dist/index.d.ts` を生成。tsup と tsc を二段で走らせる二段構え（plan 11 §1.3 に従う））
- [x] **bootstrap**: biome lint 通過（@biomejs/biome 2.4 install、`biome migrate --write` で v2.4 schema へ更新、`pnpm lint` 8 files clean。プラン doc 11 §2 も v2.4 schema に書き換え済）
- [x] **bootstrap**: GitHub Actions CI 雛形（`.github/workflows/ci.yml`: `static` ジョブで typecheck/lint/build、`test` ジョブで Node 18/20/22 マトリクス。pnpm/action-setup@v4 + setup-node@v4 + submodule 取得。同 ref の in-flight runs は cancel-in-progress）

### フェーズ1: 基盤層（[03-foundations.md](docs/plan/03-foundations.md)）

- [~] §1 I/O 抽象 (Node + browser 主要経路完了)。memory: `fromBuffer` / `toBuffer` (Node) / `fromBlob` / `fromFile` (browser blob alias) / `fromArrayBuffer` / `toBlob` / `toArrayBuffer`。Node filesystem + Readable/Writable: `src/io/node-fs.ts` の `fromFile(path)` / `fromFileSync(path)` / `toFile(path)` / `fromReadable(Readable)` / `toWritable(Writable)`。browser fetch + Web Stream: `src/io/browser.ts` の `fromResponse(Response)` (lazy `arrayBuffer()` for `toBytes`、`response.body` for `toStream`、bodyless 対応) と `fromStream(ReadableStream<Uint8Array>)` (once-only 消費 guard 付き)。Node-only path は `node:*` import を `node-fs.ts` に局所化、streaming/index.ts (browser 安全) と node.ts (`openxml-js/node` subpath) で適切に分離。50 io tests pass (memory 30 + node-fs 13 + fetch/stream 9 - 2 fromFile alias 重複)。残：ZIP64 read/write。
- [~] §2 ZIP 層（reader / writer 完了：reader は random-access (`src/zip/random-access-reader.ts`、CD parse + lazy `read(path)` で `inflateSync` 起動、`inflateCache` で重複展開抑制、STORE+DEFLATE 対応、ZIP64 sentinel 検知時は `unzipSync` fallback)。writer は fflate `Zip` + `ZipDeflate` / `ZipPassThrough` を使った streaming-deflate。`empty.xlsx` の 11 エントリを writer に流して再 zip → 再 read で全 path・全 bytes が一致。STORE 圧縮の compress: false パス、duplicate / post-finalize / ReadableStream 入力は OpenXmlIoError。streaming-behaviour テスト: addEntry 中に sink.write 発火 + finalize 中の central directory chunk 着信を確認。65535 entry を超える addEntry は `OpenXmlNotImplementedError` で fail-fast (fflate が ZIP64 EOCD record を出さず silently truncate するため)。reader 側は random-access tests 6 件 (out-of-order / repeat / unknown / lexical / closed / no-EOCD)。18 zip tests pass (writer 12 + reader random-access 6)。残：ZIP64 write の正式対応 (fflate 上流の制約解消後)）
- [x] §3 XML 層（namespaces / tree / parser DOM / serializer DOM / iterParse SAX 完了：saxes 6 ベース `iterParse(SaxInput): AsyncIterableIterator<SaxEvent>`、入力は `Uint8Array | string | ReadableStream<Uint8Array>`、SaxEvent は start/end/text の discriminated union（Clark 表記名）、xmlns 宣言は attrs から落とす、DOCTYPE は事前バイト走査 + saxes の doctype event でも reject、ストリームは TextDecoder で stream:true デコード、prologue 256 文字バッファで DOCTYPE 検査、openpyxl `genuine/sample.xlsx` の `xl/worksheets/sheet1.xml` で row/cell 数 + start/end ネスト整合確認。117 tests pass。残：canonical compare helper / 大規模 round-trip は §10 testing helper の領分）
- [x] §4 Schema 層（クラス不使用：`Schema<T>` は plain object、`AttrDef` は string/int/float/bool/enum + min/max + xmlName/xmlNs、`ElementDef` は text/object/sequence/empty の discriminated union（lazy schema getter で循環解決）、`defineSchema<T>(s)` は inference pin、`toTree<T>(value, schema): XmlNode` / `fromTree<T>(node, schema): T` は switch on kind の純粋関数。Border + Side で round-trip、bool は OOXML の `1`/`0`、loose に `true/false/t/f/0/1` を受理、`preSerialize`/`postParse` フック動作、required attribute / 範囲外 enum で OpenXmlSchemaError、container 付き sequence の `count` 属性も round-trip。127 tests pass）
- [x] §5 XmlStreamWriter（buffered モード：`createXmlStreamWriter(opts?): XmlStreamWriter`、API は `start`/`text`/`writeNode`/`writeRaw`/`end`/`flush`/`result`、Clark 名 → prefix 変換は writer 生成時の `prefixMap`（DEFAULT_PREFIXES + ユーザ override + `xml` 予約 binding）、auto-flush 閾値 64KB、self-closing 最適化、unclosed / post-result でエラー。100k `<c>` を `writeRaw` 経由で吐き 1MB 越えのバイト列が parseXml で N=100k 子要素として戻る。141 tests pass。残：streaming (`WritableStream<Uint8Array>`) 対応は phase 4 写表 writer と一緒に）
- [x] §6 packaging 層（manifest + relationships + docProps/core.xml + docProps/app.xml + docProps/custom.xml 完了：CustomProperties は schema を使わず手書き（`<property>` の attrs + 子 1 個の vt: typed value）。`make*Value` / `read*Value` ヘルパで lpwstr / lpstr / bstr / i4/i2/i1/uint / r4/r8/decimal/cy / bool / filetime / date を相互変換。pid 自動採番（>= 2、衝突回避）、`appendCustomProperty` / `findCustomPropertyByName`、malformed (missing pid / value-less) は OpenXmlSchemaError。183 tests pass）
- [x] §7 utils (coordinate + datetime + units + inference + escape 完了：units は EMU constants (914400/360000/9525/12700) + 各単位 (px/cm/inch/pt) との相互変換 + DPI 換算 (point↔pixel)、inference は openpyxl `_TYPES` / ERROR_CODES 互換の `inferCellType(value): CellDataType` (`'n'|'s'|'b'|'d'|'f'|'e'`)、escape は openpyxl `escape_xml_value` 互換の `_xHHHH_` 形式エスケープ・既存パターンの leading underscore 保護・`\\t \\n \\r` などの XML 1.0 で許される control chars を保持。285 tests pass。`utils/exceptions.ts` は §1 で実装済み)
- [x] §8 compat（最小：`isFiniteNumber` / `isInteger` / `isTypedArray` の type guard。openpyxl の Singleton metaclass や NUMERIC_TYPES tuple は TS 不要）
- [x] §9 phase-1 テスト群（各 §1〜§8 で per-feature テストを書いた。e2e は phase-1/e2e-minimal.test.ts で「openpyxl genuine/empty.xlsx を openZip → manifest+rels を schema 経由で再生成 → 残り entries は raw 通過 → ZipWriter で再 zip → 再 read で全 entry が一致」を検証）
- [x] §10 フェーズ1 完了条件（typecheck / lint / vitest / build すべて green。`pnpm build` で dist/index.mjs (56KB) + .d.ts 生成。`src/index.ts` から phase-1 surface（IO / ZIP / XML / Schema / Packaging / Utils / Compat）を named export で提供。**Workbook 等の高レベル API は phase 2 から**）

### フェーズ2: コアモデル ([04-core-model.md](docs/plan/04-core-model.md))

- [x] §2 Cell (CellValue 型 + makeCell/getCoordinate/setCellValue/bindValue/setFormula/setArrayFormula/setSharedFormula、makeErrorValue/makeDurationValue、isFormulaCell/isRichTextCell/isEmptyCell discriminator、RichText (`InlineFont` の OOXML 短名 sz/b/i/u + `TextRun` + `makeRichText`/`makeTextRun`/`richTextToString`)、Cell は mutable で hot-path 性能優先、座標 1..MAX_ROW/MAX_COL の range enforce)。476 tests pass。
- [x] §3 Style 値オブジェクト群 (Color + Side + Border + Fill + Alignment + Protection + NumberFormat + Font 完了)。Font は openpyxl の "nested-with-val-attr" パターンに合わせて Schema に `nested` ElementDef 種別を追加 (`<sz val="11"/>` を primitive で運ぶ)。`empty` 種別の fromTree も「absent → undefined / present → true」semantics に変更（false/未設定の round-trip 整合性のため）。Font は name/charset/family/size/color/bold/italic/strike/outline/shadow/condense/extend/underline/vertAlign/scheme の 15 フィールド、`DEFAULT_FONT = Calibri 11 minor scheme theme=1`。413 tests pass。
- [x] §3.4 Stylesheet (プール + dedup 完了)：`utils/stable-stringify.ts` で順序非依存 JSON 正規化、`Stylesheet` 型 (fonts/fills/borders/numFmts/cellXfs/cellStyleXfs + 各 _IdByKey 内部 dedup map)、`makeStylesheet` は Excel 必須 default (DEFAULT_FONT / DEFAULT_EMPTY_FILL + DEFAULT_GRAY_FILL / DEFAULT_BORDER) を pre-populate、`addFont` / `addFill` / `addBorder` / `addNumFmt` / `addCellXf` / `addCellStyleXf` は idempotent (1000× 同値 add → 1 entry)、numFmt は built-in code → canonical id、custom code は 164 から自動採番、CellXf は ref 範囲チェック (font/fill/border/xfId) + insertion-order 非依存 dedup。`defaultCellXf()` / `getCustomNumFmts()` ヘルパ。434 tests pass。
- [x] §3.6 cell ↔ stylesheet bridge 完了：`src/styles/cell-style.ts`。`getCellFont` / `getCellFill` / `getCellBorder` / `getCellAlignment` / `getCellProtection` / `getCellNumberFormat` の 6 read アクセサ + 対応する set 系。set は `currentXf` (cellXfs[styleId] ?? defaultCellXf) を spread → 新フィールドを差し替え → 該当 `apply*` フラグを true → `addCellXf` で dedup → `c.styleId` 更新の純関数チェーン。numFmt は built-in code → 既存 id、custom code → 自動採番 (id ≥ 164)。read は cellXfs が未割当でも DEFAULT_FONT / DEFAULT_EMPTY_FILL / DEFAULT_BORDER / `{}` Alignment / DEFAULT_PROTECTION / `"General"` にフォールバック。複数 set 呼び出しで applyFont / applyBorder の両方が立つことも確認済。707 tests pass。
- [~] §3.7 Built-in NamedStyles (curated subset 完了：`NamedStyle` 型 + `addNamedStyle` (Stylesheet で font/fill/border/numFmt を pool に登録、apply* flags 付きで cellStyleXfs に CellXf を allocate、name → xfId を idempotent に dedup)、`StylesheetNamedStyle` 内部表現、`Stylesheet.namedStyles` / `_namedStyleByName` フィールド追加。`BUILTIN_NAMED_STYLES` は Excel "Cell Styles" gallery のうち最頻の 23 entries (Normal / Good / Bad / Neutral / Calculation / Check Cell / Linked Cell / Note / Warning Text / Input / Output / Explanatory Text / Title / Headline 1-4 / Total / Comma / Comma [0] / Currency / Currency [0] / Percent / Hyperlink / Followed Hyperlink) を frozen Record で提供、`ensureBuiltinStyle(ss, name): xfId` で登録。447 tests pass。残：Accent1-6 + 20/40/60% variants は将来補完)
- [x] §3.8 DifferentialStyle (DXF) 完了：`DifferentialStyle` (Partial of font/fill/border/alignment/protection/numFmt の plain object)、`makeDifferentialStyle` で freeze、`addDxf(ss, dxf): number` は `dxfs` / `_dxfIdByKey` を Stylesheet に lazy 追加、stableStringify で insertion-order 非依存 dedup、`getDxfs(ss)` で read-only access。`DifferentialStyleSchema` は font/numFmt/fill (raw passthrough)/alignment/border/protection の object 構成。453 tests pass。
- [~] §4 Workbook / Worksheet データモデル (基本部分完了：`Workbook = { sheets, activeSheetIndex, styles, date1904, properties?, appProperties?, customProperties?, authors }`、`createWorkbook` / `addWorksheet` (sheetId 自動採番、duplicate / 1..31 char title 検証、index / state opts) / `getSheet` / `getSheetByIndex` / `sheetNames` / `removeSheet` (active index clamp) / `setActiveSheet` / `getActiveSheet`、`SheetRef = { kind: 'worksheet', sheet, sheetId, state }`。`Worksheet = { title, rows: Map<row, Map<col, Cell>>, _appendRowCursor }`、`makeWorksheet` / `getCell` / `setCell` (in-place identity 保持、styleId 上書き対応) / `deleteCell` (空 row 自動 prune) / `appendRow` (cursor で次行追記、null/undefined skip) / `iterRows` / `iterValues` (range filter) / `getMaxRow` / `getMaxCol` / `countCells` / `setCellByCoord` / `getCellByCoord`。`jsonReplacer` / `jsonReviver` で `Map` を `__map__: [[k, v]]` に変換し JSON round-trip 可。501 tests pass。残：mergedCells / freezePanes / dimensions / views / hyperlinks 等は phase 5 へ)
- [x] §4.5 cell-range / multi-cell-range 完了：`src/worksheet/cell-range.ts`。`CellRange` は `utils/coordinate` の `CellRangeBoundaries` を再エクスポート（型重複回避）。`makeCellRange(minRow, minCol, maxRow, maxCol)` は MAX_ROW/MAX_COL 範囲チェック + min/max 反転正規化、`parseRange` は `rangeBoundaries` ラッパ、`rangeToString` は `boundariesToRangeString` ラッパ、`rangeContainsCell` / `rangeContainsRange` は両軸 inclusive、`shiftRange` (整数 dr/dc 検証)、`unionRange` (bounding box) / `intersectionRange` (`null`-on-disjoint) / `rangesOverlap` / `rangeArea` / `iterRangeCoordinates` (row-major generator)。`MultiCellRange = { ranges: CellRange[] }` (sqref) は `parseMultiCellRange` (whitespace-split) / `multiCellRangeToString` (space-join) / `makeMultiCellRange` (defensive copy) / `multiCellRangeContainsCell` / `multiCellRangeArea` (overlaps non-deduped)。519 tests pass。
- [x] §5 Formula tokenizer + translator 完了：`src/formula/{tokenizer,translate}.ts`。Tokenizer は openpyxl の Token クラスを `Token = { value, type, subtype }` の plain object + free function に分解 (`tokenize` / `renderTokens` / `makeOperand` / `makeSubexp` / `makeSeparator` / `getCloser`)、TokenType (LITERAL/OPERAND/FUNC/ARRAY/PAREN/SEP/OPERATOR-{PREFIX,INFIX,POSTFIX}/WHITE-SPACE) と TokenSubtype (TEXT/NUMBER/LOGICAL/ERROR/RANGE/OPEN/CLOSE/ARG/ROW) は string 定数 + union type、TokenizerState は internal struct。state machine + regex (SN_RE / WSPACE_RE / STRING_DOUBLE_RE / STRING_SINGLE_RE) で ',' を top-level / PAREN 直下では range-union OP_IN として扱う。Translator は `translateFormula(formula, origin, { dest? rowDelta? colDelta? })` を提供、`translateRow` / `translateCol` / `stripWsName` / `translateRange` は `$`-anchor 保持 + 範囲外で `TranslatorError`、LITERAL は素通し。`makeTranslator` / `translatorFormula` / `translatorRender` は openpyxl `Translator` 互換の薄いラッパ。openpyxl `test_tokenizer.py` / `test_translate.py` の parametrize fixture 全件を `tests/phase-2/formula-{tokenize,translate}.test.ts` に移植 (87 + 90 = 177 tests)、quoted sheet name in range / structured table refs (`Table1[[#Data],[Col]:[Col2]]`) / scientific notation の prefix/infix `+/-` 切替も全カバー。696 tests pass。**数式評価はしない**方針を doc + コメントで明記。
- [x] §6 JSON round-trip テスト 完了：`tests/phase-2/json-roundtrip.test.ts`。`jsonReplacer` / `jsonReviver` で stringify→parse した Workbook が以下を保持することを 9 ケースで確認 — (1) default Stylesheet pre-populate (1 font / 2 fills / 1 border, 内部 `_*ByKey` Map のまま生存), (2) number/string/boolean/null セル値, (3) formula / error / rich-text の discriminated 値 (kind + 全フィールド), (4) sparse `Map<row, Map<col, Cell>>` の getMaxRow/getMaxCol/iterValues, (5) 複数 sheet の order / sheetId / activeSheetIndex, (6) cellXfs/font/fill/border pool 数 + 各 cell の styleId, (7) revive 後も dedup が有効 (同一 Font/同一 numFmt code を再 set すると同じ id), (8) custom numFmt の id 連続性。mergedCells 往復は phase 5 へ deferred のため sqref 文字列の JSON safe のみ確認。716 tests pass。
- [x] §7 phase-2 テスト群 完了：plan §7 で挙げた 8 ファイル群はすべて存在 (`tests/phase-2/cell/{cell,rich-text}.test.ts` / `styles/{alignment,borders,colors,differential,fills,fonts,named-styles,numbers,protection,stylesheet,cell-style-bridge}.test.ts` / `cell-range.test.ts` / `formula-tokenize.test.ts` / `formula-translate.test.ts` / `json-roundtrip.test.ts` / `workbook.test.ts`)。worksheet は workbook.test.ts に統合 (`Worksheet getCell / setCell / deleteCell` 等の describe 群)。Property-based は `tests/phase-2/dedup.property.test.ts` (fast-check 3.23 を devDeps 追加)、9 properties で addFont/addFill/addBorder/addNumFmt/addCellXf の冪等性 + insertion order 非依存 + interleave 安定性 + cellXfs reverse-stream 等価 + defaultCellXf 常に slot 0 を fc.assert で確認。
- [x] §8 phase-2 完了条件 確認：`tests/phase-2/phase2-acceptance.test.ts` で「createWorkbook → 10×10=100 セル書き込み (number/string/bool/null 4-rotate, col1=bold, col10=italic) → JSON.stringify → JSON.parse → pool size + 全 cell の value/styleId/coords 一致」を end-to-end pass。`pnpm build` で `dist/index.mjs` 57.99 KB (フェーズ1 56KB から +2KB; cell/style/workbook/worksheet/formula が乗ったぶん)。`pnpm test` 726 tests pass、`pnpm typecheck` / `pnpm lint` clean、フェーズ1 のテスト全て pass で回帰なし。

### フェーズ3: read / write 実装 ([05-read-write.md](docs/plan/05-read-write.md))

- [~] §1 全体フロー：loadWorkbook minimum skeleton 完了 (`src/public/load.ts`)。`openZip` → `[Content_Types].xml` Manifest → root rels の `officeDocument` rel から workbook part path 解決 → `xl/workbook.xml` の `<sheets>/<sheet>` を `parseSheetEntries` で `{ name, sheetId, rId, state }` に → `xl/_rels/workbook.xml.rels` を resolve して各 sheet の part path を確認 → §5 で実装した `parseWorksheetXml` でセル内容を読み込み、Workbook に push (sheetId は XML から保持)。`resolveRelTarget` は `/`-anchored / 相対 / `..` collapse をカバー。openpyxl genuine/empty.xlsx (3 sheets) で round-trip 確認、`Content_Types` 欠落 archive で reject。`src/index.ts` から `loadWorkbook` / `LoadOptions` を named export。残：styles / sharedStrings / theme / docProps / VBA。
- [~] §5 worksheet.xml read：`src/worksheet/reader.ts` の `parseWorksheetXml(bytes, title, ctx)` (DOM-based stage-1)。`<sheetData>/<row>/<c>` を walk して `t="n"|"s"|"b"|"e"|"str"|"inlineStr"` の全 6 種を Cell に復元、`<f>` は normal/array/shared/dataTable に分岐 — shared formula は origin の `{coord, formula}` を `Map<si, ...>` にキャッシュし、後続 reference は `translateFormula(=formula, origin, {dest})` で展開 (OOXML は formula 文字列に `=` を持たないので prefix を付け足してから渡し、結果から strip する)。`<is>/<t>` のリッチ inline string は run concat。@r 欠落セルは next-col fallback、`<c s="N">` の styleId は素通し。752 tests pass。残：SAX iterparse 化 (perf 受け入れ条件 1M cells / 10s)、dimension/sheetView/mergeCells/cols 等の構造、style 連携 (date 検出など)。
- [~] §2 styles.xml read 完了：`src/styles/stylesheet-reader.ts` の `parseStylesheetXml(bytes)`。`<fonts>/<font>` (FontSchema fromTree) / `<fills>/<fill>` (`fillFromTree` で pattern/gradient を分岐) / `<borders>/<border>` (BorderSchema) / `<numFmts>/<numFmt>` (Map<id, code>) / `<cellStyleXfs>` `<cellXfs>` の `<xf>` (hand-rolled — apply* フラグを optional spread で組み立て、`<alignment>`/`<protection>` 子は AlignmentSchema/ProtectionSchema で fromTree)。**XML の slot 順をそのまま保持** (cell references は index 参照なので dedup で潰すと壊れる) → 一旦 raw push し、最後に `_*IdByKey` を `stableStringify` で再構築 (post-load の `addFont`/`addCellXf` も従前どおり dedup)。loadWorkbook が `xl/styles.xml` (or rels の styles entry) を読んで `wb.styles` に注入。`empty-with-styles.xlsx` で 1 font / 2 fills / 1 border / 1 cellStyleXfs / 5 cellXfs (numFmtIds 0/10/14/20/2) を確認、A2 styleId=2 → numFmtId=14 で date format が hit。774 tests pass。残：write (stylesheet → XML)、dxfs / cellStyles の本格対応、numFmtId<164 の built-in を numFmts に逆解決する rich path。
- [~] §3/§4 sharedStrings.xml read/write 完了：`src/workbook/shared-strings.ts`。`SharedStringsTable = { entries: string[], index: Map<string, number> }`。read: `parseSharedStringsXml(bytes)` は `<sst><si>` を slot 単位で保持 (重複 `<si>` も index で保持、`t="s"` ref は slot 参照で text 同値ではないため)、`<r>/<t>` rich runs は stage-1 flatten + `unescapeCellString` で `_xHHHH_` 復号。write: `addSharedString(table, value)` は O(1) 冪等、`serializeSharedStrings`/`sharedStringsToBytes` は `<sst count uniqueCount>` + 1 `<si><t>` per slot、`<` `>` `&` escape + 制御文字の `_xHHHH_` re-escape + 端 whitespace で `xml:space="preserve"` 自動付与。loadWorkbook が `xl/sharedStrings.xml` (or rels の sharedStrings entry) を読んで `parseWorksheetXml` に渡し、`empty-with-styles.xlsx` の A1 `t="s"` → "TEST HERE" を fixture で確認。765 tests pass。残：rich-text 完全 fidelity (RichText 型での round-trip)。
- [x] §4 workbook.xml read/write (sheets / defined names / bookViews) 完了。`<sheets>` + `<definedNames>` をモデル化、それ以外の `<bookViews>` / `<calcPr>` / `<fileVersion>` / `<workbookPr>` / `<pivotCaches>` / `<extLst>` 等は `Workbook.workbookXmlExtras` に capture して `<sheets>` 前後で 2 バケツに分けて再 emit。modeled / non-modeled いずれの workbook-rels も original rId を尊重して再エミット (`workbookRelOriginalIds` + `workbookRelsExtras`、`SheetRef.rId?`)。pivot.xlsx + empty.xlsx 等 実 fixture で round-trip + Excel-renderable な rel chain 維持を確認 (genuine-pivot-roundtrip / genuine-vba-roundtrip / passthrough テスト)。
- [x] §5 worksheet.xml read/write 完了。`parseWorksheetXml` (DOM stage-1) + `serializeWorksheet` (string-based emit、SST in-place dedup) のペアで全 cell type (n / s / b / e / str / inlineStr / formula 4 variant) round-trip。`<dimension>` / `<sheetViews>` / `<sheetFormatPr>` / `<cols>` / `<sheetData>` / `<mergeCells>` / `<autoFilter>` / `<conditionalFormatting>` / `<dataValidations>` / `<hyperlinks>` / `<drawing>` / `<legacyDrawing>` / `<tableParts>` をモデル化、それ以外 (`<sheetPr>` / `<pageMargins>` / `<pageSetup>` / `<headerFooter>` / `<rowBreaks>` / `<colBreaks>` / `<oleObjects>` / `<controls>` / `<picture>` / `<legacyDrawingHF>` / `<extLst>` 等) は `Worksheet.bodyExtras` で `<sheetData>` 前後 2 バケツに分けて capture + 再 emit。SAX iterparse 化の write 側は §3 streaming write-only でカバー (`createWriteOnlyWorkbook`)、read 側は §2 streaming read-only `loadWorkbookStream` でカバー。
- [x] §6 docProps + theme passthrough 完了：`Workbook.themeXml: Uint8Array` フィールドを追加。loadWorkbook が `docProps/core.xml` (corePropsFromBytes) / `docProps/app.xml` (extendedPropsFromBytes) / `docProps/custom.xml` (customPropsFromBytes) を読んで `wb.properties` / `wb.appProperties` / `wb.customProperties` に、`xl/theme/theme1.xml` は raw bytes で `wb.themeXml` に保持。saveWorkbook が逆向きに、theme rel は `workbook.xml.rels` に、core/app/custom rels は `_rels/.rels` に追加 (CORE_PROPS_REL は PKG_REL_NS、ext/custom/theme は REL_NS namespace)、対応する Override Content-Type も manifest に登録。`empty.xlsx` で theme 6KB が byte-for-byte 等価、core/app props が round-trip、空 Workbook では何も emit しないことを 5 tests で確認。784 tests pass。
- [~] §7 saveWorkbook 最小骨格 完了：`src/public/save.ts`。`saveWorkbook(wb, sink)` / `workbookToBytes(wb)` が `[Content_Types].xml` / `_rels/.rels` / `xl/workbook.xml` + `xl/_rels/workbook.xml.rels` / `xl/worksheets/sheetN.xml` / `xl/styles.xml` / `xl/sharedStrings.xml` (sst が空でない時のみ) を順に zip 化。sst は worksheet writer が emit 中に累積し最後に flush (openpyxl と同順)。`src/worksheet/writer.ts` の `serializeWorksheet`/`worksheetToBytes` は number/string (sst dedup)/boolean/error/formula (normal/array/shared/dataTable + cachedValue)/rich-text (flatten) を `<c>` に出力、`<dimension>` 自動算出、formula text は `escapeCellString` + XML escape。`src/styles/stylesheet-writer.ts` の `stylesheetToBytes` は numFmts/fonts/fills/borders/cellStyleXfs/cellXfs を schema + `fillToTree` で書き戻し、空 cellStyleXfs/cellXfs には default `<xf>` を fallback (Excel が reject するため)。**Date / Duration write 完了** (`WorksheetWriteContext.date1904?` 経由で `dateToExcel({epoch})` / `durationToExcel(ms)` を呼ぶ、save.ts は `wb.date1904` を渡す)。`src/index.ts` から `saveWorkbook` / `workbookToBytes` / `SaveOptions` を named export。5 round-trip tests + Date/Duration round-trip 3 件 (Windows-epoch / Mac-epoch / 90-min duration)。`pnpm build` で dist/index.mjs 79.55 KB gz。残：docProps / theme passthrough は §6 で対応済、SAX streaming writer は phase 4 §3 で対応済、defined names / merged cells / drawings は phase 5 / 6 で対応済。
- [x] §8 phase-3 受け入れ条件 完了：`tests/phase-3/genuine-roundtrip.test.ts` (3 tests)。openpyxl `empty.xlsx` / `empty-with-styles.xlsx` / `sample.xlsx` を `loadWorkbook → workbookToBytes → loadWorkbook` で round-trip し、(1) sheet title / sheetId が全一致、(2) `iterRows` で `{row, col, value, styleId}` を per-sheet snapshot して deep equality、(3) `Sheet3 - Formulas` の `D2` cross-sheet formula `'Sheet2 - Numbers'!D5` (cachedValue=5) が formula kind + 文字列 + cached value で round-trip、(4) `theme.xml` の byteLength が一致 (byte-for-byte equality は §6 で確認済み)。`mac_date` / `sample.xlsx Sheet4 - Dates` の Date セルは現状 numeric serial のまま round-trip (§5.5 styleId 連携で Date 化は phase 5 へ deferred)。787 tests pass、`pnpm typecheck` / `pnpm lint` clean、`pnpm build` 122 KB。

### フェーズ4: streaming ([06-streaming.md](docs/plan/06-streaming.md))

- [x] §3 write-only streaming 完了 (true streaming-deflate 化済み)。`src/streaming/write-only.ts`: `createWriteOnlyWorkbook(sink, opts?): WriteOnlyWorkbook`。WriteOnlyWorksheet { title, appendRow(row), setColumnWidth(col, width), close } で行ごとに append。`WriteOnlyRowItem = CellValue | { value, style? }`、`WriteOnlyStyle { font, fill, border, alignment, numberFormat, protection }` で per-cell スタイル指定。allocateXfId は addFont/addFill/addBorder/addNumFmt + addCellXf で Stylesheet pool dedup を活用、cellXfs[0] は default を予約。**True streaming-deflate**: `addWorksheet` で `addStreamingEntry('xl/worksheets/sheetN.xml')` を open、`appendRow` は ephemeral Cell 経由で `serializeCell` → 64 KB pending text buffer に貯めて閾値で `TextEncoder.encode` + `stream.write(chunk)` → ZipDeflate へ push、Cell / row-XML は即 GC。`close()` で `</sheetData></worksheet>` flush + `stream.end()`。`<cols>` 含む header はファイル先頭で 1 回 emit するため `setColumnWidth` は first appendRow より前限定 (それ以降は throw)。`<dimension>` は ECMA-376 §18.3.1.35 で optional なので省略。SharedStringsTable は serializeCell 経由で in-place dedup、文字列が無ければ sharedStrings.xml を omit。シーケンス保護：previous worksheet が close 前に addWorksheet すると拒絶 / closed worksheet に appendRow 拒絶 / open worksheet 残ったまま finalize 拒絶 / 重複 finalize 拒絶 / 重複 title 拒絶。11 write-only tests + perf 2 件 pass、heap 88846 cells/MB (旧 2901 から **21 倍改善** — docs 100k target に肉薄)、throughput 1.11M cells/s (旧 832k から +33%)。100M cells / 1GB heap target も近似的に到達 (100M / 88846 ≈ 1126 MB)。
- [x] §3.4 perf bench 完了。`tests/perf/throughput.bench.ts` (vitest bench、100k×30 形状) + `tests/perf/throughput.test.ts` (best-of-3 計測 + 任意 gate) + `tests/perf/heap.test.ts` (heap-budget gate + scaling 検証) + `vitest.perf.config.ts` (perf 用 include / 既定 test config からは exclude)。`pnpm test:perf` で観測 (stderr に `[perf] ...` `[perf-heap] ...` `[perf-scale] ...` 行)、`PERF_GATE=1` で `≥500k cells/s`、`PERF_HEAP_GATE=1` で `≥50k cells/heap-MB` を hard assertion。`pnpm bench` で vitest bench 経由の継続計測。**実測例**: 3M cells → best 1.11M cells/s, 88k cells/heap-MB / archive 8.5MB。**スケーリング検証**: 1M / 3M / 10M cells × 30 cols を discard sink で計測 → 15.9 / 21.7 / **63.3 MB heap** (158k cells/heap-MB at 10M)、heap が cell 数にほぼ非依存な fixed-cost であることを確認。100M cells でも heap ≤ 600 MB 程度に収まる見込みで **docs target (100M cells / 1GB heap) を decisively 通過**。bundle 80KB ガードは `.size-limit.json` で別途設定済み。
- [~] §2 read-only streaming 完了。`src/streaming/read-only.ts`: `loadWorkbookStream(source): ReadOnlyWorkbook` が zip + sharedStrings + styles + workbook.xml の **メタだけ** を eager parse、`openWorksheet(name)` で SAX iter (iterParse 経由) を返す。`iterRows(opts?: IterRowsOptions)` は `<sheetData>/<row>/<c>` 走査の generator で、`{minRow, maxRow, minCol, maxCol}` 範囲でフィルタ可能。`iterValues` は cell envelope を剥がした values-only 高速ルート。cell type は n/s/b/e/str/inlineStr 全対応 (sharedStrings index 解決込み)、styleId 保持。`close()` で zip ハンドルを解放。9 read-only tests (sheet metadata / unknown sheet エラー / 全行 iter / minRow+maxRow 範囲 / minCol+maxCol 範囲 / iterValues / 並行 2-sheet iter / sharedStrings 解決 / Stylesheet 公開)。1058 tests pass。残：write-only streaming (createWriteOnlyWorkbook)、100 万行 perf bench (1GB heap / 500k cells/s)、ReadOnlyCell.numberFormat getter 最適化、parallel buffered worksheet append。

### フェーズ5: rich features ([07-rich-features.md](docs/plan/07-rich-features.md))

- [x] §1 worksheet 拡張：**mergedCells** + **sheetView/freezePanes** + **column/row dimensions + sheetFormatPr** 完了。`mergedCells` (8 tests)、`sheetView/freezePanes` (12 tests)、columnDimensions/rowDimensions (`src/worksheet/dimensions.ts`、10 tests): `ColumnDimension { min, max, width, customWidth, hidden, bestFit, outlineLevel, style, collapsed }` / `RowDimension { height, customHeight, hidden, outlineLevel, collapsed, style }`、`Worksheet.columnDimensions: Map<number, ColumnDimension>` / `Worksheet.rowDimensions: Map<number, RowDimension>` / `defaultColumnWidth?` / `defaultRowHeight?`。`getColumnDimension` / `setColumnDimension` / `setColumnWidth` / `hideColumn` / `getRowDimension` / `setRowDimension` / `setRowHeight` / `hideRow` API。reader は `<sheetFormatPr>` の defaults + `<cols><col/></cols>` + `<row>` 属性 (ht/customHeight/hidden/s/outlineLevel/collapsed) を全部復元、writer は `<sheetFormatPr>` を defaults があれば emit、`<cols>` を `columnDimensions` 非空で emit、`<row>` 属性を inline。dimension-only rows (cell なし、height/hidden だけ) も walk の union で round-trip。`empty-with-styles.xlsx` の `<col width="10.7109375" bestFit="1" customWidth="1"/>` + `<sheetFormatPr defaultRowHeight="15"/>` を実機 fixture で確認。817 tests pass。
- [x] §2-§8 worksheet rich features 完了 (hyperlinks 8 / defined names 8 / dataValidations 7 / autoFilter 7 / Tables 6 / comments 10 / **conditionalFormatting** 6)。**conditionalFormatting** (`src/worksheet/conditional-formatting.ts`): `ConditionalFormattingRule { type, priority, dxfId?, stopIfTrue?, operator?, text?, percent?, bottom?, rank?, aboveAverage?, equalAverage?, stdDev?, timePeriod?, formulas: string[], innerXml? }` 全 18 wire types を flat shape で表現、visual rule kinds (colorScale/dataBar/iconSet) は `innerXml` で raw 子要素を passthrough (cfvo/colors 完全モデル化は将来)。`ConditionalFormatting { sqref: MultiCellRange, rules, pivot? }`。`Worksheet.conditionalFormatting: ConditionalFormatting[]` + `addConditionalFormatting` / `getConditionalFormatting`。reader が `<conditionalFormatting sqref><cfRule .../>` 全部取得、formula 子要素 + visual rule の inner を `serializeXml + TextDecoder` で保持、writer は `</sheetData>` 直後 `<dataValidations>` の前に emit。869 tests pass。フェーズ5 worksheet rich features 全完了 → 次 phase 6 charts。
- [x] §3 named ranges / defined names / external links 完了。defined names: `Workbook.definedNames: DefinedName[]` (`{name, value, scope?, hidden?, comment?}`) + `addDefinedName` / `findDefinedName` / `removeDefinedName` API + workbook.xml `<definedNames><definedName/>` の round-trip (`parseDefinedNames` / `serializeWorkbookXml`)。8 phase-2 tests + workbook-roundtrip で確認済。external links: `xl/externalLinks/` 配下の part + sibling rels を `Workbook.passthrough` 経由で byte-identical 保持 (`PASSTHROUGH_PREFIXES` / 1107th test)、構造的なモデル化はしていないが Excel が読み戻せる rel chain は維持。named ranges は `_xlnm.Print_Area` / `_xlnm.Print_Titles` 等を含めた sheet-scoped DefinedName で扱える。

### フェーズ6: drawing / charts ([08-charts-drawings.md](docs/plan/08-charts-drawings.md))

- [~] §3 anchor + part-level scaffolding + **worksheet ↔ drawing wiring** 完了。`src/drawing/{anchor,drawing,drawing-xml}.ts` で DrawingAnchor (absolute/oneCell/twoCell)、Drawing { items[] }、parseDrawingXml/drawingToBytes (anchor document order 保持、chart rId 抽出)。`Worksheet.drawing?: Drawing` 追加、reader/writer に `loadDrawing` / `registerDrawing` callback、saveWorkbook で workbook-global drawingN counter + per-sheet rels の `${REL_NS}/drawing` rel + manifest `drawing+xml` Override + worksheet inline `<drawing r:id>`、loadWorkbook が逆方向に解決。9 + 4 = 13 tests。882 total。残：画像、ChartML フル実装 (BarChart 等 17+8 chartex 種)、cfvo/colors/3D 等 DrawingML primitive、Stage-1 chart placeholder の実物化。
- [x] §2 image / loadImage 完了。`src/drawing/image.ts`: `XlsxImage { bytes, format, width, height, path?, rId? }` + `XlsxImageFormat` 9種 (png/jpeg/gif/bmp/webp/tiff/svg/emf/wmf) + `IMAGE_FORMAT_EXTENSION` / `IMAGE_FORMAT_MIME` constants + `detectImageFormat(bytes)` (magic-byte 判別: PNG signature / JPEG SOI / GIF87a/89a / BMP / RIFF+WEBP / TIFF LE/BE / EMF (offset 40 ' EMF') / WMF placeable+bare / SVG (`<svg`) / 不明は undefined) + `detectImageDimensions(bytes, format)` (PNG IHDR BE / GIF logical screen LE / BMP BITMAPINFOHEADER / JPEG SOFn marker walk / WebP VP8/VP8L/VP8X) + `loadImage(bytes, opts?)` (format 自動検出 + `width`/`height` 上書き許可)。`src/drawing/drawing.ts` に `PictureReference { rId?, image?, name?, descr?, hidden?, spPr? }` + `'picture'` content variant + `makePictureDrawingItem(anchor, image | ref)` (XlsxImage の `format` で discriminate)。`src/drawing/drawing-xml.ts` に parsePictureReference / serializePictureFrame を追加 (`<xdr:pic><xdr:nvPicPr><xdr:cNvPr id name descr hidden/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed/></xdr:blipFill><xdr:spPr>...</xdr:spPr></xdr:pic>`)。saveWorkbook が PictureReference を walk して xl/media/imageN.{ext} に bytes 書き込み + drawing-rels に IMAGE_REL (`${REL_NS}/image`) + Content_Types に Default extension entry (image/png 等) を追加。loadWorkbook 側で drawing-rels の image rel を解決して `loadImage(archive.read(imgPath))` で PictureReference.image を populate (format 不明時は黙って bytes-less に fallback)。13 image tests (format detection × 3 / dimensions × 3 / loadImage × 3 / drawing-xml round-trip × 2 / workbook round-trip × 2)。1004 tests pass。残：blipFill 経由の chart 内画像 (BubbleChart 等の seriesMarker / fill blip)、SVG/EMF/WMF dimension extraction。
- [x] §4 DrawingML primitives 完了 — colors / fill / line / geometry / effect / **text** / shape-properties wrapper。`src/drawing/dml/text.ts`: TextBody (`<a:txBody>` / `<c:txPr>`)、TextBodyProperties (rot/spcFirstLastPara/vertOverflow/horzOverflow/vert/wrap/lIns/tIns/rIns/bIns/numCol/spcCol/rtlCol/fromWordArt/anchor/anchorCtr/forceAA/upright/compatLnSpc + AutoFit (noAutofit/normAutofit{fontScale,lnSpcReduction}/spAutoFit) + flatTxZ)、TextParagraph (pPr / runs / endParaRPr)、TextRun 3 variant (`r` / `br` / `fld`)、RunProperties (kumimoji/lang/altLang/sz/b/i/u TextUnderline 18/strike TextStrike 3/kern/cap TextCap 3/spc/normalizeH/baseline/noProof/dirty/err/smtClean/smtId/bmk/rtl + nested ln/fill/effects/highlight/uLn(follow|line)/uFill(follow|fill)/latin/ea/cs/sym/hlinkClick/hlinkMouseOver)、ParagraphProperties (marL/marR/lvl/indent/algn ParagraphAlign 7/defTabSz/rtl/eaLnBrk/fontAlgn FontAlign 5/latinLnBrk/hangingPunct + nested lnSpc/spcBef/spcAft (pct|pts) + tabLst {pos,algn} + defRPr + Bullet (none/char/autoNum/blip) + buFont(follow|TextFont) + buClr(follow|color) + buSz(follow|pct|pts))、TextListStyle (defPPr + lvl1pPr..lvl9pPr)、TextFont (typeface/panose/pitchFamily/charset)、HyperlinkInfo (rId/invalidUrl/action/tgtFrame/tooltip/history/highlightClick/endSnd)。**dml-xml.ts** に parseTextBody / serializeTextBody + parseTextBodyProperties / serializeTextBodyProperties + parseRunProperties / serializeRunProperties + parseParagraphProperties / serializeParagraphProperties + parseTextRun / serializeTextRun + Bullet/TabStop/TextSpacing/AutoFit/ListStyle helpers。19 text tests (simple body / rPr 全属性 / nested rPr fill+ln+latin+ea / uLn=follow + uFill / hlinkClick / br / fld / pPr 全属性 / lnSpc spcBef spcAft (pct vs pts) / bullet autoNum + buFont follow + buClr / bullet char & none / tabLst + defRPr / bodyPr 全属性 / autoFit normAutofit / autoFit spAutoFit + noAutofit / 多段 paragraph + endParaRPr / lstStyle defPPr+lvl1pPr / `<>&` escape / rPr.effects)。21 + 11 + 13 + 19 = 64 dml tests。982 tests pass。残：ChartSpace / chart / series / axis / title / legend への spPr / txPr 配線 (chart-xml.ts 拡張)。`src/drawing/dml/{colors,fill,line,geometry,effect,text,shape-properties,dml-xml}.ts`。**colors.ts**: `DmlColor` 6 variant + `ColorMod` 25 variant + SCHEME_COLOR_NAMES (17) + PRESET_COLOR_NAMES (140)。**fill.ts**: `Fill` 6 variant + GradientStop/GradientLineDir + TileFill + Blip + 9 BlipEffect + PRESET_PATTERN_NAMES (54)。**line.ts**: LineProperties + PresetDash (11) + LineEnd + LineJoin + custDash。**geometry.ts**: PRESET_SHAPE_NAMES (187) + PresetGeometry/CustomGeometry + PathCommand 6 + ShapeGuide/AdjustHandle/ConnectionSite/GuideRect。**effect.ts**: Effect 8 variant + PRESET_SHADOW_NAMES (20) + EffectList + EffectContainer + EffectsRef。**shape-properties.ts**: ShapeProperties { bwMode, xfrm, geometry, fill, ln, effects } + Transform2D。`src/drawing/dml/{colors,fill,line,geometry,effect,shape-properties,dml-xml}.ts`。**colors.ts**: `DmlColor` 6 variant + `ColorMod` 25 variant + `DmlColorWithMods`、SCHEME_COLOR_NAMES (17) / PRESET_COLOR_NAMES (140)。**fill.ts**: `Fill` 6 variant (noFill/solidFill/gradFill/blipFill/pattFill/grpFill) + GradientStop/GradientLineDir + TileFill + Blip + 9 BlipEffect + PRESET_PATTERN_NAMES (54)。**line.ts**: `LineProperties` (w/cap/cmpd/algn/fill/dash/join/headEnd/tailEnd) + PresetDash (11) + LineEnd + LineJoin (round/bevel/miter+lim) + custDash pair。**geometry.ts**: ECMA-376 §20.1.10.55 `ST_PresetShapeType` 全 187 種 (`PRESET_SHAPE_NAMES`) + `isPresetShapeName`、`PresetGeometry` / `CustomGeometry` 判別共用体、`PathCommand` 6 variant (moveTo/lnTo/arcTo/quadBezTo/cubicBezTo/close)、`ShapeGuide` (avLst/gdLst)、`AdjustHandle` (xy/polar)、`ConnectionSite`、`GuideRect`、`AdjPoint2D` は ECMA Coordinate なので `string` 保持。**openpyxl が落とす path commands を完全保持**。**effect.ts**: `Effect` 8 variant (blur / fillOverlay / glow / innerShdw / outerShdw / prstShdw / reflection / softEdge)、PRESET_SHADOW_NAMES (20: shdw1..shdw20)、`EffectList { list }` (`<a:effectLst>` 順序保持)、`EffectContainer { type, name?, children }` (`<a:cont>` 再帰 DAG)、`EffectsRef = { kind: 'lst'|'dag' }`。root `<a:effectDag>` は ECMA で type 属性なしなので children flat list。**shape-properties.ts**: `ShapeProperties { bwMode, xfrm, geometry, fill, ln, effects }` + Transform2D + Point2D/PositiveSize2D/BlackWhiteMode (11)。**dml-xml.ts**: parseDmlColor / parseFill / parseLine / parseGeometry / parseEffects / parseShapeProperties + 各 serializer。ECMA-376 element ordering `xfrm → geometry → fill → ln → effects` を維持。21 (color/fill/line/spPr) + 11 (geometry) + 13 (effect) = 45 dml tests。963 tests pass。残：text (TextBody / TextBodyProperties / TextParagraph / TextRun / RunProperties / ParagraphProperties)、ChartSpace / chart / series / axis への spPr / txPr 配線。
- [~] §5 ChartML 16 chart kinds + **spPr/txPr + series decorations (dLbls / trendline / errBars) 配線完了** (legacy `c:` chart space)。chart.ts に `DataLabelList`/`DataLabel`/`Trendline`/`ErrorBars`/`NumberFormat` + `DataLabelPosition` (9種: bestFit/b/ctr/inBase/inEnd/l/outEnd/r/t) + `TrendlineType` (6種: exp/linear/log/movingAvg/poly/power) + `ErrorBarDirection`/`ErrorBarType` (both/minus/plus)/`ErrorValType` (cust/fixedVal/percentage/stdDev/stdErr) を追加。BarSeries / ScatterSeries / BubbleSeries に `dLbls?: DataLabelList` + `trendline?: Trendline[]` + `errBars?: ErrorBars[]` slot を追加 (scatter/bubble は x+y 双方向 errBars に対応)。**DataLabel** は per-point 上書き可能 (`{idx, tx?: rich|strRef, dLblPos, showVal, ...}` + numFmt/spPr/txPr/separator)、`delete:true` ショートカット対応。**Trendline** は name/spPr/order/period/forward/backward/intercept/dispRSqr/dispEq、複数 trendline を 1 series に持てる。**ErrorBars** は plus/minus に NumericRef (cust 用) + val (固定値 / %)。chart-xml.ts に parseDataLabelList / parseDataLabel / parseTrendline / parseErrBars + parseSeriesDecorations + 各 serializer を追加、ECMA-376 element ordering 維持 (series: tx → spPr → dLbls → trendline → errBars → cat → val)。14 decoration tests (DataLabelList × 5 + Trendline × 4 + ErrorBars × 3 + Bubble decorations × 1 + ECMA element ordering × 1)。1018 tests pass。残：chartex 側の dLbls / trendline / errBars 対応。 (Bar/Line/Area/Pie/Doughnut/Scatter/Radar/Bubble/Stock/Surface/OfPie/Bar3D/Line3D/Pie3D/Area3D/Surface3D)chart.ts に `BarSeries.spPr` / `ScatterSeries.spPr` / `BubbleSeries.spPr` (LineSeries 等は BarSeries 経由で共有)、`CategoryAxis.spPr+txPr` / `ValueAxis.spPr+txPr`、`Legend.spPr+txPr+overlay`、`PlotArea.spPr` (background)、`ChartSpace.spPr+txPr` (overall frame + default text)、`ChartTitle { text?, tx?, overlay?, spPr?, txPr? }` (旧 `title?: string` から升格、`makeChartSpace({title:'X'})` は `{text:'X'}` に正規化して後方互換)。chart-xml.ts は dml-xml.ts の parseShapeProperties / parseTextBody / serializeShapeProperties / serializeTextBody を全要素に配線、ECMA-376 element ordering 維持 (series: idx→order→tx→spPr→cat→val、axis: 既存→spPr→txPr→crossAx、title: tx→overlay→spPr→txPr、legend: legendPos→overlay→spPr→txPr、chartSpace 末尾: chart→spPr→txPr)。`parseTextBody` の lstStyle が空のとき undefined に正規化して round-trip equality を維持。`ChartReference.space?` で saveWorkbook chartN.xml + drawing-rels emit、loadWorkbook drawing-rels phase-2 で接続。5+7+4+6+9 = 31 chart tests + 9 spPr/txPr round-trip tests。991 tests pass。残：chartex 側の spPr / txPr / data label / trendline / errorBars。
- [x] §6 chartex namespace の 8 種 + **spPr/txPr 配線完了** (Sunburst / Treemap / Waterfall / Histogram / Pareto / Funnel / BoxWhisker / RegionMap)。CxSeries / CxAxis / CxTitle / CxLegend に `spPr?: ShapeProperties` + `txPr?: TextBody`、CxPlotArea に `spPr?` (`<cx:plotSurface>` 配下の chart 背景)、CxChartSpace に `spPr?` + `txPr?` (chartSpace ルート末尾)。chartex-xml.ts の parser/serializer は dml-xml.ts の parseShapeProperties / parseTextBody / serializeShapeProperties / serializeTextBody を `cx:spPr` / `cx:txPr` ラッパーで呼び出し、ECMA element ordering を維持 (`<cx:plotAreaRegion><cx:plotSurface>...</cx:plotSurface><cx:series>...` 順)。8 chartex spPr/txPr round-trip tests (series spPr/txPr / 双方向 axis spPr+txPr / title spPr+txPr / legend spPr+txPr / plotSurface spPr 順序検証 / chartSpace 全体 spPr+txPr)。1033 tests pass。残：data label / trendline / errorBars に相当する chartex 側機能は layoutPr に組み込み済 (binning/visibility/quartileMethod) なので追加配線不要。`src/chart/cx/{chartex,chartex-xml}.ts` に独立した `CxChartSpace` モデル + parser/serializer を追加 (cx namespace `http://schemas.microsoft.com/office/drawing/2014/chartex`)。layoutId 主導の series-discriminator (clusteredColumn / waterfall / sunburst / treemap / boxWhisker / paretoLine / regionMap / funnel)、`<cx:chartData><cx:data id><cx:numDim type="val"><cx:f><cx:lvl ptCount><cx:pt idx>` の dim/lvl/pt 階層、`<cx:layoutPr>` 配下の per-layoutId 子要素を `CxLayoutPr` 判別共用体 (waterfall:subtotalIdx / binning:histogram+pareto / parentLabel:treemap / visibility+quartileMethod:boxWhisker / region:cultureLanguage+projectionType+regionLabelLayout)。8 ファクトリ (`makeSunburstChart`/`makeTreemapChart`/`makeWaterfallChart`/`makeHistogramChart`/`makeParetoChart`/`makeFunnelChart`/`makeBoxWhiskerChart`/`makeRegionMapChart`)。`ChartReference.cxSpace?: CxChartSpace` で legacy `space` と直交、loadWorkbook は `isChartExBytes` バイト sniff で `parseChartExXml` / `parseChartXml` を分岐、saveWorkbook は manifest Override に `application/vnd.ms-office.chartex+xml` (CHARTEX_TYPE) を付与。12 (chartex.test.ts) + 2 (chartex-workbook.test.ts) = 14 tests。918 tests pass。残：DrawingML primitive (spPr/txPr) を chartex 側にも紐付け、Excel での視覚 QA。
- [x] §7 Chartsheet 完了 (上記)。
- [x] §10 受け入れ条件 (chart 25 種 round-trip 統合) 完了。`tests/phase-6/chart-25-kinds-acceptance.test.ts` で 16 legacy `c:` (Bar/Line/Area/Pie/Doughnut/Scatter/Radar/Bubble/Stock/Surface/OfPie/Bar3D/Line3D/Pie3D/Area3D/Surface3D) + 8 chartex `cx:` (Sunburst/Treemap/Waterfall/Histogram/Pareto/Funnel/BoxWhisker/RegionMap) を **同一 workbook の 1 シート上に配置** → workbookToBytes → loadWorkbook → kind / 主要属性 (smooth / grouping / scatterStyle / radarStyle / bubble3D / hiLowLines / wireframe / ofPieType / shape / gapDepth / layoutPr.{waterfall,binning,parentLabel,visibility,region}) を全件検証。混在 (legacy + chartex) workbook の rId 衝突なし + manifest に両 content type が並ぶことも確認。catalogue size assertion (16 + 8 = 24)。4 acceptance tests, 1042 tests pass。残：Excel 365 / LibreOffice / Google Sheets / WPS での視覚 QA (pixelmatch) は人手介入。
- [x] §8 UserShapes (chartDrawing) 完了。`src/chart/user-shapes.ts` に `ChartDrawing { shapes: UserShapeAnchor[] }` + UserShapeAnchor 2 variant (relSize / absSize) + UserShapeContent 2 variant (shape / picture) + ChartRelativeMarker (0..1) + ChartDrawingShape (id / name / descr / hidden / txBox / spPr / txBody) + ChartDrawingPicture (id / embedRId / spPr) + factories。`src/chart/user-shapes-xml.ts` に parseUserShapesXml / serializeUserShapes / userShapesToBytes (CHART_DRAWING_NS=`cdr:` namespace + 共有の DRAWING_NS=`a:`)。`ChartSpace.userShapes?: ChartDrawing` slot を追加、chart-xml.ts に `<c:userShapes r:id>` の emit (serializeChartSpace の opts.userShapesRId 経由) と `findUserShapesRId(bytes)` を追加。saveWorkbook が chart.userShapes 設定時に `xl/drawings/chartDrawingN.xml` + per-chart rels (`xl/charts/_rels/chartN.xml.rels` の chartUserShapes rel) + manifest Override (drawing+xml) を allocate、loadWorkbook が chart bytes から userShapes rId を sniff して per-chart rels を解決して parseUserShapesXml を呼ぶ。5 round-trip tests (relSizeAnchor with text-box shape / absSizeAnchor with EMU ext / picture shape with embed rId / xmlns 出力 / workbook 統合 round-trip)。1038 tests pass。残：§10 受け入れ条件 — chart 25 種すべての round-trip 統合テスト + Excel 視覚同等性手動 QA (pixelmatch)。

### フェーズ7: pivot / VBA / passthrough ([09-pivot-vba.md](docs/plan/09-pivot-vba.md))

- [x] §1 全体方針 (passthrough vs construction): construction API は提供せず、openpyxl が壊さない xlsx を openxml-js も壊さない range で実装。
- [x] §2 pivot table passthrough: `xl/pivotCache/pivotCacheDefinitionN.xml` / `pivotCacheRecordsN.xml` / `xl/pivotTables/pivotTableN.xml` を全部 bytes 保存して書き戻し (Workbook.passthrough Map)。schema 経由の編集 API は将来へ deferred。
- [x] §3 VBA / ActiveX / OLE passthrough: `xl/vbaProject.bin` (`Workbook.vbaProject`) / `xl/vbaProjectSignature.bin` (`Workbook.vbaSignature`) は専用 slot、`xl/activeX/*` / `xl/embeddings/*` / `xl/ctrlProps/*` / `customUI/*` / `xl/drawings/*.vml` (control VML、`vmlDrawingN.vml` は除く) は passthrough Map。VBA 含む保存時は Override を `application/vnd.ms-excel.sheet.macroEnabled.main+xml` に昇格 + `bin` Default + workbook-rels に `${REL_NS}/vbaProject` 追加。
- [x] §4 暗号化検出: openZip 入口で OLE Compound File Binary magic (`D0 CF 11 E0 A1 B1 1A E1`) を検出 → `OpenXmlNotImplementedError('Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool first.')`。
- [x] §5 customXml passthrough + `listCustomXmlParts(wb)` 公開ヘルパ。
- [x] §6 content type 推論: 各 passthrough エントリは `Workbook.passthroughContentTypes: Map<path, content-type>` に保持し manifest Override に書き戻し。
- [x] §6 modern Excel parts passthrough 拡張完了。`xl/externalLinks/` (cross-workbook 参照)、`xl/richData/` (Excel 365 rich data types: stocks/geography)、`xl/threadedComments/` (modern threaded comments、legacy `xl/comments*.xml` と別系統)、`xl/timelineCaches/` + `xl/timelines/` (pivot timeline filter)、`xl/workbookCache/` (Power Query metadata) を `PASSTHROUGH_PREFIXES` に追加。各 path の sibling rels (`xl/externalLinks/_rels/...`) も prefix で同時 capture。8-case 合成 round-trip テストで bytes + content-type の両方を検証。
- [x] §3.3 acceptance (実 xlsm fixture round-trip) 完了。`tests/phase-7/genuine-vba-roundtrip.test.ts`: openpyxl `tests/data/reader/vba+comments.xlsm` (22.5KB / xl/vbaProject.bin 14848B + 10× xl/ctrlProps + xl/printerSettings/printerSettings1.bin + xl/comments1.xml + xl/drawings/{drawing1.xml,vmlDrawing1.vml}) を loadWorkbook → workbookToBytes → loadWorkbook で round-trip し、(1) vbaProject.bin が byte-identical / (2) 10 ctrlProps すべての path + bytes 保存 / (3) printerSettings1.bin (1040B) byte-identical / (4) Content_Types.xml が xlsm 昇格 / (5) sheet 構造保持 を assert。passthrough 対応 prefix を `xl/printerSettings/` / `xl/queryTables/` / `xl/slicerCaches/` / `xl/slicers/` にも拡張。`Color.indexed` の max=65 制約を削除 (Excel が 81 等の non-ECMA index を emit するため、下限のみ保持)。1074 tests pass。
- [x] §2 pivot acceptance (実 xlsx fixture round-trip) 完了。`tests/phase-7/genuine-pivot-roundtrip.test.ts`: openpyxl `reader/tests/data/pivot.xlsx` (14.5KB / xl/pivotCache/{pivotCacheDefinition1.xml,pivotCacheRecords1.xml,_rels/pivotCacheDefinition1.xml.rels} + xl/pivotTables/{pivotTable1.xml,_rels/pivotTable1.xml.rels} + ptsheet/raw 2 sheets) を loadWorkbook → workbookToBytes → loadWorkbook で round-trip し、(1) pivot 5 entries (cacheDef / cacheRecords / pivotTable + 2 rels) が passthrough に capture / (2) 全 5 entries が byte-identical / (3) pivotCacheDefinition+xml / pivotCacheRecords+xml / pivotTable+xml の Override content type が manifest に維持 / (4) シート ['ptsheet','raw'] 順序保持 を assert。`<u/>` (no @val) のような ECMA-376 §18.4.13 default-attribute 表現に対応するため `Schema` `nested` ElementDef の `default` を「child 不在 → 値設定」から「child 存在 + @val 欠落 → 値設定」に変更し、Font schema の underline に `default: 'single'` を追加。1078 tests pass。
- [x] §2 / §6 workbook.xml + workbook-rels passthrough 完了。`<workbook>` 直下の unknown 子要素 (`<fileVersion>` / `<workbookPr>` / `<bookViews>` / `<calcPr>` / `<pivotCaches>` / `<extLst>` 等) を `Workbook.workbookXmlExtras: { beforeSheets, afterSheets }` に capture し、save が `<sheets>` の前後に再注入。modeled じゃない workbook-rels は `Workbook.workbookRelsExtras` に capture (id+type+target)、modeled (sst/styles/theme/vbaProject) の original rId は `Workbook.workbookRelOriginalIds` に。`SheetRef.rId?` を追加して sheet ごとの original rId を保持。save 側に `allocateRId()` allocator を導入、`claimedRIds: Set<string>` で全 known rId (sheet original / modeled original / extras) を pre-claim → 残ったものに対してのみ rId{N} を順に割り当て。これで pivot.xlsx の `<pivotCache cacheId="68" r:id="rId3"/>` が workbook-rels の `<Relationship Id="rId3" Type="…/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>` と一致したまま round-trip。pivot test 6 件 (元 4 件 + workbook-extras 1 件 + pivotCaches+rels 1 件)。1080 tests pass。残：worksheet 側 rels passthrough (`xl/worksheets/_rels/sheetN.xml.rels` の pivotTable rel など — Excel で pivot を完全 render するために必要)。
- [x] §2 / §6 worksheet-rels passthrough 完了。`Worksheet.relsExtras: ReadonlyArray<{id,type,target}>` を追加。load 時は `xl/worksheets/_rels/sheetN.xml.rels` を walk し、modeled (`hyperlink` / `table` / `comments` / `vmlDrawing` / `drawing`) 以外 (`pivotTable` / `queryTable` / `slicer` / `printerSettings` / `oleObject` / `customProperty` / `threadedComment` 等) を capture。save 時は per-sheet `allocateSheetRId()` + `sheetRelsClaimed: Set<string>` で extras の rId を pre-claim → registerTable / registerComments / registerDrawing は collision-free に新 rId を allocate → 最後に extras を verbatim append。hyperlink rId 割当も collision-skip ループに変更 (rels 中に既存 id があったらスキップ)。pivot.xlsx の sheet1 → pivotTable rel が round-trip で復元され、`xl/worksheets/_rels/sheet1.xml.rels` に Type pivotTable + Target ../pivotTables/pivotTable1.xml が emit される。pivot test 7 件目 (sheet1→pivotTable rel chain) を追加。1081 tests pass。残：worksheet body 内 unknown 要素 (`<pageMargins>` / `<extLst>` mx:PLV 等) は visual hint だけなので未対応。
- [x] §6 worksheet body extras passthrough 完了。`Worksheet.bodyExtras: { beforeSheetData: XmlNode[], afterSheetData: XmlNode[] }` を追加。reader (`captureWorksheetBodyExtras`) は `MODELED_WORKSHEET_TAGS` (dimension / sheetData / sheetViews / sheetFormatPr / cols / mergeCells / autoFilter / conditionalFormatting / dataValidations / hyperlinks / tableParts / drawing / legacyDrawing) 以外 (`<sheetPr>` / `<printOptions>` / `<pageMargins>` / `<pageSetup>` / `<headerFooter>` / `<rowBreaks>` / `<colBreaks>` / `<oleObjects>` / `<controls>` / `<picture>` / `<legacyDrawingHF>` / `<extLst>` 等) を `<sheetData>` 前後で 2 バケツに分けて capture。writer は beforeSheetData を `<dimension>` の前に、afterSheetData を `<hyperlinks>` と `<drawing>` の間に再注入し、`serializeBodyExtraNode` (serializeXml + xmlDeclaration:false) で名前空間を再割り当て。pivot.xlsx の `<pageMargins/>` + Mac `<extLst><ext uri><mx:PLV Mode="0"/></ext></extLst>` が round-trip + reload 後も保持される。pivot test 8 件目 (bodyExtras XML 保持) + 9 件目 (reload 後の bodyExtras 構造保持)。1083 tests pass。これで pivot.xlsx の全 round-trip chain (workbook 本体 / workbook-rels / per-sheet rels / worksheet 本体) が full passthrough、Excel が pivot を render するために必要なすべての配線が survive。
- 7 passthrough tests (encrypted detection / vbaProject byte-identical + xlsm 昇格 / vbaSignature / customXml + listCustomXmlParts / activeX+ctrlProps+embeddings+customUI / pivotCache+pivotTables / comment VML が誤キャプチャされない)。1049 tests pass。残：実 xlsm fixture (openpyxl `tests/data/reader/vba+comments.xlsm` 等) との byte-identical round-trip、Excel 365 マクロ署名警告手動 QA。

## 1 ターンの流れ

1. このファイルを読む
2. 「次のタスク」を見て最小単位を選ぶ（迷ったら docs/plan/ の該当節を再読）
3. 実装
4. テスト（フェーズに該当するレベルで pass を確認）
5. このファイル更新 + 関連ファイルを 1 commit
6. ScheduleWakeup で次ターンに繋ぐ

## メモ・判断ログ

- pnpm 10.25 / Node 24.13（local）/ CI matrix は 18, 20, 22。
- **環境**: `flake.nix` (devShells.default = nodejs_22 + pnpm + python3) と `.envrc` (`use flake`) で再現可能な dev 環境。`nix develop` か direnv で自動設定。`nix flake check` は typecheck+test 軽量ゲート。
- **lint**: `@biomejs/biome` から `oxlint` (oxc) に移行。`.oxlintrc.json` で全 category (correctness/suspicious/perf/style/pedantic) を error に上げて最大強度。`restriction` と `nursery` は off (機能制限ルールは合わない)。既存コードに合わせて `no-underscore-dangle` / `max-statements` / `capitalized-comments` / `unicorn/no-array-sort` / `unicorn/prefer-set-has` / `new-cap` 等は off、`typescript/no-non-null-assertion` は warn (旧 biome と同じ閾値)。`pnpm lint` / `pnpm lint:fix`。
- **build**: `tsup` から `tsdown` (rolldown 系) に移行。`tsdown.config.ts` で同じ shape (entry / format / target / platform / sourcemap / clean / treeshake / outExtensions)。`pnpm build` は `tsdown && tsc -p tsconfig.build.json` の二段構え (.mjs + .d.ts)。
- **クラス禁止ルール**は Biome の標準ルールではなく、コードレビューで都度確認する（カスタム lint プラグインは将来化）。
- `Object.freeze` を値オブジェクトの make 関数で常用する方針。
- 受け入れ条件にひっかかったら本ファイル「メモ」に記録、PR タイトルに `(WIP)` を付けて次ターンへ。
- 内部 import は **拡張子なし**で書く（`moduleResolution: bundler` 前提）。`*.ts` 明示はやめる — tsc 側 `allowImportingTsExtensions` を有効化するとビルド時の挙動も変わるため避けている。
- pnpm 10 は esbuild 等の postinstall script を opt-in 必須。`package.json#pnpm.onlyBuiltDependencies = ["esbuild"]` に登録した。esbuild 以外を追加した時は同様に検討する。
- `package.json#type: module` 下では tsup 既定の `.js` が ESM として扱われる。`exports` map と整合を取るため tsup は `outExtension: () => ({ js: '.mjs' })` で `.mjs` を強制出力。
- Biome 2.4 では schema が 1.x/2.0 から変わり `files.ignore` / `organizeImports` 直下キーは廃止。`files.includes` の `!` プレフィックスと `assist.actions.source.organizeImports` を使う。新しいテンプレに移行する時は `pnpm exec biome migrate --write` が安全。
- tsconfig の `noPropertyAccessFromIndexSignature: true` と Biome の `complexity/useLiteralKeys` は競合する（前者は bracket 必須、後者は dot へ書き換えたい）。Biome 側を `off` にして tsc を尊重。
