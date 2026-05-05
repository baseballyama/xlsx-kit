# 13. openpyxl の枠を越えて Excel 全機能をサポートするためのロードマップ

本リポジトリは openpyxl の TypeScript port として始まったが、最終目標は
**Excel が「ネイティブに」できる全機能** を openxml-js で読み書きできる
ことである。openpyxl 自身がサポートしない機能 (Excel 365 の動的配列、
LAMBDA、threaded comments の編集、Power Query の構造アクセスなど) を含めて
網羅する必要がある。

本ドキュメントは、その差分を「カテゴリ → 必要な作業の概要 → ECMA-376 /
[MS-XLSX] 参照 → 推定スプリント数」の形でまとめる。実装順は依存関係と
ユースケースの広さで判断する。

> 凡例:
> - **Sprint** は 1 名フルタイム週単位
> - **Status** は本リポジトリの現状: ✅ 完了 / 🟡 部分対応 / ❌ 未着手 / 🔒 外部 blocked

---

## A. 既に passthrough 経由で round-trip するが編集 API がない (高優先度)

| # | 機能 | Status | 必要な作業 | Sprint |
|---|------|--------|-----------|-------|
| A1 | **Pivot Table 編集** (`xl/pivotTables/`, `xl/pivotCache/`) | 🟡 passthrough 維持 | (1) `<pivotCacheDefinition>` / `<pivotCacheRecords>` / `<pivotTableDefinition>` の schema 化、(2) PivotTable オブジェクトモデル (rowFields / colFields / pageFields / dataFields / filters)、(3) 既存 sheet データ → pivotCache を build する API、(4) cache の再計算ヒューリスティック (`refreshOnLoad`)。ECMA-376 §18.10 + [MS-XLSX] §2.4.227 | 4–6 |
| A2 | **Threaded Comments 編集** (`xl/threadedComments/`, `xl/persons/`) | 🟡 passthrough | `<threadedComment>` schema、`Person` registry、parent/child reply tree、`legacyComment` (yellow note) との 1:1 ペアリング。MC-IGNORABLE で legacy をフォールバックさせる仕組みが必要 | 2 |
| A3 | **Slicer / Timeline 編集** (`xl/slicers/`, `xl/slicerCaches/`, `xl/timelines/`, `xl/timelineCaches/`) | 🟡 passthrough | 各 `<slicer>` / `<timeline>` の schema + Pivot との接続 (slicerCache → pivotCache)。使い勝手は Pivot 編集 (A1) に依存 | 2–3 |
| A4 | **External Links 編集** (`xl/externalLinks/`) | 🟡 passthrough | `<externalLink>` schema、外部 workbook 参照の `[1]Sheet1!$A$1` 形式の formula token、`<externalBook>` cache | 2 |
| A5 | **Custom XML 部品の構造アクセス** | 🟡 passthrough | `customXml/itemN.xml` + `itemPropsN.xml` の schema + `<dataRecoveryPart>` 経由の content controls。Word/PowerPoint 共有 | 1 |
| A6 | **VBA プロジェクト** (`xl/vbaProject.bin`) | 🔒 byte-identical 保持のみ | バイナリ OLE Compound File parser → `dir` / `_VBA_PROJECT` stream の解析 → モジュール単位の編集。署名 (`xl/vbaProjectSignature.bin`) の再生成は X.509 証明書チェーン要 | 8–12 |
| A7 | **OLE Embeddings / ActiveX** (`xl/embeddings/`, `xl/activeX/`) | 🟡 passthrough | OLE Compound File parser 共有 (A6 と同基盤)、`<oleObject>` / `<control>` の schema、worksheet inline `<oleObjects>` / `<controls>` セクション | 4 |
| A8 | **Power Pivot Data Model** (`xl/model/`) | 🟡 passthrough | `xl/model/item.data` + `xl/model/item.xml` の解析。Microsoft 内部仕様で公式 schema が無いため reverse engineering が必要 | 6+ |
| A9 | **Power Query** (`xl/queryTables/`, `xl/connections.xml`, `xl/workbookCache/`) | 🟡 passthrough | M クエリの schema + 外部データソース定義。Power Query Formula Language の token 化が必要 | 4+ |

---

## B. ECMA-376 で定義済だが本リポジトリ未着手 (中優先度)

| # | 機能 | Status | 必要な作業 | Sprint |
|---|------|--------|-----------|-------|
| B1 | **Sparklines** (in-cell mini charts) | ❌ | `<x14:sparklineGroups>` (`mc:AlternateContent`)。data range / type (line/column/winLoss) / styling。worksheet の extLst に格納 | 2 |
| B2 | **Dynamic Arrays / SPILL** | ❌ | Excel 365 の `<f t="array" aca="1" ca="1">` + `<metadata>` cell. ARRAY-aware formula tokenizer。動的配列 SPILL エラーの round-trip。ECMA-376 第6版以降 | 3 |
| B3 | **LAMBDA / LET / 名前付き関数** | ❌ | `<definedName>` に lambda 式を入れる仕様。formula tokenizer が `LAMBDA(x, x*2)` 構文を扱える必要 | 2 |
| B4 | **数式評価 (formula evaluation engine)** | ❌ | 全数式の評価系 — 範囲計算、数学関数 200+、文字列関数、論理関数、参照関数 (INDIRECT 等)、配列関数。**最大規模の作業項目** | 20+ |
| B5 | **Workbook / Sheet protection** | 🟡 部分対応 | `<sheetProtection>` 全 16 boolean lock flag + 4 password-hash field を typed API + round-trip 完了 (`src/worksheet/protection.ts`、`tests/phase-5/sheet-protection.test.ts`、e2e 27 が `makeSheetProtection` 経由)。残：パスワード hash 計算 helper (legacy + SHA-512)、`<workbookProtection>` の typed API | 2 |
| B6 | **Print Settings** (rowBreaks / colBreaks / pageSetup / headerFooter / printOptions) | 🟡 worksheet body extras で passthrough | 各要素の schema + 編集 API。breakManually / orientation / paperSize / scale / fitToWidth/Height | 2 |
| B7 | **Sheet view** 拡張 (sheetPr / customSheetViews / scenarios / dataConsolidate) | 🟡 部分対応 | sheetPr 全体 (codeName / tabColor / outlinePr / pageSetUpPr 含む全 9 attr + 3 child) は `src/worksheet/properties.ts` で typed API + round-trip 完了 (e2e 27 が `makeSheetProperties` 経由で tab color を着色)。残：customSheetViews / scenarios / dataConsolidate | 1–2 |
| B8 | **Cell Watches / Ignored Errors** | ✅ | `<cellWatches>` / `<ignoredErrors>` schema (`src/worksheet/errors.ts`、reader/writer wired、helper API + round-trip tests in `tests/phase-5/errors.test.ts`) | 0.5 |
| B9 | **Web Publish Items / Custom Properties (worksheet level)** | ❌ | `<webPublishItems>` / `<customProperties>` schema | 0.5 |
| B10 | **Phonetic / EaList (East Asian features)** | ❌ | `<phoneticPr>` (ふりがな)、`<rPh>` per cell。日本語 / 中国語 workbook では普通に出現 | 1 |
| B11 | **Filter Database / Outline (group/ungroup rows)** | ✅ | `<row outlineLevel>` 読み書き可、`<sheetFormatPr outlineLevelRow/Col + customHeight/zeroHeight/thickTop/thickBottom/baseColWidth>` 全 round-trip。outlineLevelRow/Col は明示値があればそれを優先、無ければ row/columnDimensions から auto-compute。`tests/phase-5/sheet-format-pr.test.ts` 4 件 | 0.5 |

---

## C. ECMA-376 拡張 (Microsoft 名前空間 `x14` / `x15` / `xr` / `xr2` / etc.)

| # | 機能 | Status | 必要な作業 | Sprint |
|---|------|--------|-----------|-------|
| C1 | **Conditional Formatting 拡張** (gradient cfvo / iconSet 拡張) | 🟡 stage-1 (innerXml 通過) | x14 namespace の `<x14:conditionalFormatting>` / `<x14:dataBar>` etc. を schema 化 | 2 |
| C2 | **Data Validation 拡張** (custom errorMessage XML、prompt) | 🟡 部分 | x14 拡張 attrs (`x14:dataValidation`) | 0.5 |
| C3 | **Pivot V2 (xr3 / xr10)** | 🟡 passthrough | x15:pivotTablesItems / xr3:pivotCacheDefinition の schema | 2 |
| C4 | **Excel 365 Cell Metadata** (`xl/metadata.xml`) | 🟡 passthrough | metadata types (string / numeric / arrayRef / valueMetadata) の schema、`<c cm="N">` の意味付け | 1 |
| C5 | **Rich Data Types** (Stocks / Geography) (`xl/richData/`) | 🟡 passthrough | rdRichValueTypes / rdRichValue / rdRichValueStructure。DataModel に近い | 3 |
| C6 | **Sensitivity Labels / DRM** | ❌ | `customXml/MIP*` の schema、暗号化と組み合わせた IRM フロー | 1 |
| C7 | **Inking** (`xl/ink/`) | ❌ | `<inkAnnotations>` / `<x14:inkPath>`。Surface Pen 入力データ | 1 |
| C8 | **3D Geometry / 3D Charts (chartex 拡張)** | 🟡 部分 | chartex 8 種は対応済みだが、3D bar / 3D pie の visual properties (rotation / depth / perspective) は ChartML legacy 経由のみ | 2 |

---

## D. 暗号化 / 署名 / 完全性

| # | 機能 | Status | 必要な作業 | Sprint |
|---|------|--------|-----------|-------|
| D1 | **暗号化 xlsx の復号** | 🔒 検出のみ | OLE Compound File parser、ECMA-376 Part 4 §3 (Office crypto API)、AES-128 / AES-256 / Standard / Agile encryption の実装、パスワード入力 UX | 4–6 |
| D2 | **暗号化 xlsx の生成** | ❌ | D1 の inverse + 鍵管理 | 3 |
| D3 | **デジタル署名** (`_xmlsignatures/`) | ❌ | XML-DSIG + OOXML signature manifest (`_xmlsignatures/sig*.xml`)、X.509 証明書管理 | 4 |
| D4 | **DRM / IRM** | ❌ | Microsoft AD RMS protocol、暗号化と密結合 | 4+ |

---

## E. 性能 / I/O / 互換性 (ライブラリ品質)

| # | 機能 | Status | 必要な作業 | Sprint |
|---|------|--------|-----------|-------|
| E1 | **ZIP64 write** | 🔒 fflate 上流 blocked | fflate に PR (ZIP64 EOCD record の emit) **または** 自前 zip writer (大変だが直球) | 3 |
| E2 | **Multi-process / Web Worker streaming write** | ❌ | parallel sheet build + worker pool で deflate を分散 | 3 |
| E3 | **真のランダムアクセス reader** (Excel-style "open large file in <1s") | 🟡 部分 | row-offset index は実装済 (`tail-of-sheet` で 3415x speedup)。残：sharedStrings の遅延 lookup、文字列が多い workbook で sst を on-demand index に変更 | 2 |
| E4 | **Browser Web Worker サンドボックス** | ❌ | 既存 streaming API を `OffscreenCanvas` 風に worker 内で実行できる API ラッパー | 1 |
| E5 | **Excel for Mac / LibreOffice / Google Sheets / WPS / Numbers での視覚 QA** | 🔒 人手 | pixelmatch 比較を CI に組む、または手動チェックリスト (本リポの `tests/e2e/` がその下準備) | 継続的 |
| E6 | **Excel-computed cachedValue の検証** | ❌ | 数式評価エンジン (B4) ができれば、書き出し時に自前で計算した値と Excel が再計算する値の一致を比較できる | B4 と同時 |
| E7 | **複数言語の数値書式** (`#,##0` の locale-aware separator) | ❌ | Excel は workbook の locale に応じて表示を切り替える。numFmt code 自体は locale-independent。読み込み時の `_x002A_` セパレータ等の解釈 | 1 |

---

## F. 内部品質 / 開発エルゴノミクス

| # | 項目 | Status | 必要な作業 |
|---|------|--------|-----------|
| F1 | TypeScript strict + exactOptionalPropertyTypes 全 module で完備 | ✅ | — |
| F2 | size-limit ガード (full / streaming / 各 subpath) | ✅ | サブパスが増えるごとに limit 追加 |
| F3 | typedoc 自動生成 + GitHub Pages publish | ❌ | `pnpm doc:api` script + workflow |
| F4 | 移行ガイド (`docs/migrate-from-openpyxl.md`) | ✅ | 大変更時に更新 |
| F5 | README / 公式サイト | 🟡 README あり | 公式サイト (vitepress / astro) に doc を載せる |
| F6 | 自動 fixture 検査 (load → save → load → diff archive bytes) | ✅ ad-hoc | CI で定期実行する形に組み込む |
| F7 | Mutation testing (stryker) | ❌ | 未着手。1.0 release 後に検討 |
| F8 | パフォーマンス CI 比較 (PR で perf 数値を delta 表示) | ❌ | size-limit と同じ仕組みで perf-bot |
| F9 | Bundle 構成 (`openxml-js/streaming` など) | ✅ | サブパス追加ごとに更新 |

---

## マイルストーン提案

### `1.0` (現状): openpyxl パリティ + Excel が壊さない round-trip 保証
- Phase 1-7 + 全 passthrough chain + 25 chart 種 + streaming reader/writer の acceptance
- e2e フィクスチャ視覚 QA + edge-fixture deep-assert
- Excel 365 機能の **読み書き** は passthrough 経由で全て round-trip
- **編集 API** は openpyxl と同等の範囲 (Pivot / VBA / Power Query は read-only passthrough)

### `1.1`: Threaded Comments / Slicers / Timelines の編集 API
- A2 + A3 を完成
- Excel 365 で書いたファイルが変わらず Excel 365 で開ける
- Pivot は touched-row tracking で部分編集可能

### `1.2`: Sparklines / Dynamic Arrays / LAMBDA
- B1 + B2 + B3
- Excel 365 の最新数式機能を round-trip + 編集

### `1.3`: 数式評価エンジン
- B4 — 巨大な作業。`SUM` / `AVERAGE` / `IF` / `VLOOKUP` / `INDEX/MATCH` / `XLOOKUP` / `LET` / `LAMBDA` / 配列関数を順に
- これが完成すると **「Excel ファイルを Excel 抜きで完結に処理できる」**

### `1.4`: Pivot Table / Power Query の編集 API
- A1 + A9 — Pivot を build できれば BI ダッシュボード自動生成が可能になる

### `2.0`: 暗号化 / デジタル署名 / VBA
- D1 + D2 + D3 + A6
- 企業ユースケース完備
- VBA の編集 API で macro-enabled workbook の自動生成

### `2.1+`: Power Pivot Data Model / Sensitivity Labels / Inking
- A8 + C5 + C6 + C7
- Excel が将来追加する機能に追従するための拡張点

---

## 進め方の原則

1. **passthrough を先に固める**: schema 化していない part も bytes が保たれていれば「壊さない」契約は守れる。本リポは現時点で全 part が round-trip するため、**今後の編集 API 追加は upward-compatible**。
2. **ECMA-376 を一次資料とする**: openpyxl の挙動に固執せず、ECMA-376 と Excel の実際の挙動が食い違ったら ECMA-376 + Excel emit を優先 (openpyxl 自身も同じ流儀)。
3. **fixture-driven**: 各機能の追加は openpyxl / 自作 / Microsoft 公式の Excel ファイルを round-trip 通すことを必須要件にする。`tests/e2e/` にシナリオ追加 + `tests/phase-*/genuine-edge-fixtures.test.ts` に deep-assert 追加。
4. **段階的 reveal**: openpyxl が壊さない部分から先に schema 化し、編集 API は後追い。schema 化する前に passthrough で bytes が回ることが前提。
5. **bundle 予算を破らない**: 機能を増やすときは必ず subpath で分離。`openxml-js` (full) ≤ 200 KB / `openxml-js/streaming` ≤ 80 KB / 大規模機能は `openxml-js/pivot` `openxml-js/powerquery` 等の独立サブパス。

---

## 関連リソース

- ECMA-376: <https://ecma-international.org/publications-and-standards/standards/ecma-376/>
- [MS-XLSX]: <https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/>
- Microsoft Office binary file formats: <https://learn.microsoft.com/en-us/openspecs/office_file_formats/>
- openpyxl docs: <https://openpyxl.readthedocs.io/>
- ExcelJS (TypeScript 競合実装): <https://github.com/exceljs/exceljs> — 編集 API の参考に。ライセンス互換性は要確認
- xlsx-populate / xlsx (SheetJS): <https://docs.sheetjs.com/> — 数式評価の参考
