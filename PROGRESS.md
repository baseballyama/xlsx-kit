# PROGRESS

`/loop` 自走モードでフェーズ1→7 を順に実装するための、ターン横断の状態ファイル。
**唯一の正は `docs/plan/`**。本ファイルは「いまどこまで終わったか」だけを記録する。

## カレント

- **フェーズ**: フェーズ2 (コアモデル)
- **次のタスク**: フェーズ2 §4 Workbook / Worksheet データモデル。`Workbook = { sheets: SheetRef[], styles: Stylesheet, properties, ... }` + `createWorkbook` / `addWorksheet` / `removeSheet` / `getSheet` / `setActiveSheet` / `defineName`。`Worksheet = { title, rows: Map<row, Map<col, Cell>>, columnDimensions, rowDimensions, mergedCells, views, ... }` + `getCell` / `setCell` / `appendRow` / `iterRows` / `mergeCells` / `setColumnWidth` / `setRowHeight` / `setFreezePanes`。Worksheet → Workbook back-ref (mutate 用)、JSON.stringify replacer で循環参照対策。続けて §4.5 cell-range / multi-cell-range、§5 Formula tokenizer / translator。
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

- [~] §1 I/O 抽象（メモリ経路のみ完了：`XlsxSource` / `XlsxSink` / `BufferedSinkWriter` の interface、`OpenXmlError` 階層、Node の `fromBuffer` / `toBuffer`、ブラウザの `fromBlob` / `fromFile` / `fromArrayBuffer` / `toBlob` / `toArrayBuffer`、30 tests pass。残：filesystem / Readable / Writable / Response 経路は §2 ZIP streaming と同時に）
- [~] §2 ZIP 層（reader / writer メモリ経路完了：`fflate.unzipSync` の `openZip` + `fflate.zipSync` の `createZipWriter`。`empty.xlsx` の 11 エントリを writer に流して再 zip → 再 read で全 path・全 bytes が一致。STORE 圧縮の compress: false パス、duplicate / post-finalize / ReadableStream 入力は OpenXmlIoError。47 tests pass。残：streaming reader / streaming writer / ZIP64 read/write）
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
- [ ] §3.6 cell ↔ stylesheet bridge (`getCellFont` / `setCellFont` 等の free function)
- [~] §3.7 Built-in NamedStyles (curated subset 完了：`NamedStyle` 型 + `addNamedStyle` (Stylesheet で font/fill/border/numFmt を pool に登録、apply* flags 付きで cellStyleXfs に CellXf を allocate、name → xfId を idempotent に dedup)、`StylesheetNamedStyle` 内部表現、`Stylesheet.namedStyles` / `_namedStyleByName` フィールド追加。`BUILTIN_NAMED_STYLES` は Excel "Cell Styles" gallery のうち最頻の 23 entries (Normal / Good / Bad / Neutral / Calculation / Check Cell / Linked Cell / Note / Warning Text / Input / Output / Explanatory Text / Title / Headline 1-4 / Total / Comma / Comma [0] / Currency / Currency [0] / Percent / Hyperlink / Followed Hyperlink) を frozen Record で提供、`ensureBuiltinStyle(ss, name): xfId` で登録。447 tests pass。残：Accent1-6 + 20/40/60% variants は将来補完)
- [x] §3.8 DifferentialStyle (DXF) 完了：`DifferentialStyle` (Partial of font/fill/border/alignment/protection/numFmt の plain object)、`makeDifferentialStyle` で freeze、`addDxf(ss, dxf): number` は `dxfs` / `_dxfIdByKey` を Stylesheet に lazy 追加、stableStringify で insertion-order 非依存 dedup、`getDxfs(ss)` で read-only access。`DifferentialStyleSchema` は font/numFmt/fill (raw passthrough)/alignment/border/protection の object 構成。453 tests pass。
- [ ] §4 Workbook / Worksheet データモデル
- [ ] §4.5 cell-range / multi-cell-range
- [ ] §5 Formula tokenizer + translator
- [ ] §6 JSON round-trip テスト
- [ ] §7 phase-2 テスト群
- [ ] §8 phase-2 完了条件

## 1 ターンの流れ

1. このファイルを読む
2. 「次のタスク」を見て最小単位を選ぶ（迷ったら docs/plan/ の該当節を再読）
3. 実装
4. テスト（フェーズに該当するレベルで pass を確認）
5. このファイル更新 + 関連ファイルを 1 commit
6. ScheduleWakeup で次ターンに繋ぐ

## メモ・判断ログ

- pnpm 10.25 / Node 24.13（local）/ CI matrix は 18, 20, 22。
- **クラス禁止ルール**は Biome の標準ルールではなく、コードレビューで都度確認する（カスタム lint プラグインは将来化）。
- `Object.freeze` を値オブジェクトの make 関数で常用する方針。
- 受け入れ条件にひっかかったら本ファイル「メモ」に記録、PR タイトルに `(WIP)` を付けて次ターンへ。
- 内部 import は **拡張子なし**で書く（`moduleResolution: bundler` 前提）。`*.ts` 明示はやめる — tsc 側 `allowImportingTsExtensions` を有効化するとビルド時の挙動も変わるため避けている。
- pnpm 10 は esbuild 等の postinstall script を opt-in 必須。`package.json#pnpm.onlyBuiltDependencies = ["esbuild"]` に登録した。esbuild 以外を追加した時は同様に検討する。
- `package.json#type: module` 下では tsup 既定の `.js` が ESM として扱われる。`exports` map と整合を取るため tsup は `outExtension: () => ({ js: '.mjs' })` で `.mjs` を強制出力。
- Biome 2.4 では schema が 1.x/2.0 から変わり `files.ignore` / `organizeImports` 直下キーは廃止。`files.includes` の `!` プレフィックスと `assist.actions.source.organizeImports` を使う。新しいテンプレに移行する時は `pnpm exec biome migrate --write` が安全。
- tsconfig の `noPropertyAccessFromIndexSignature: true` と Biome の `complexity/useLiteralKeys` は競合する（前者は bracket 必須、後者は dot へ書き換えたい）。Biome 側を `off` にして tsc を尊重。
