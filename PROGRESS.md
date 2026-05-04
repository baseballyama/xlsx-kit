# PROGRESS

`/loop` 自走モードでフェーズ1→7 を順に実装するための、ターン横断の状態ファイル。
**唯一の正は `docs/plan/`**。本ファイルは「いまどこまで終わったか」だけを記録する。

## カレント

- **フェーズ**: フェーズ1（基盤層）
- **次のタスク**: フェーズ1 §6 packaging 層の続き：`docProps/custom.xml` (CustomProperties)。`<property fmtid pid name>` 配下に `<vt:lpwstr>` / `<vt:i4>` / `<vt:filetime>` / `<vt:bool>` / `<vt:r8>` 等の typed value 1 個。最初は schema の sequence + 値型 1 個ずつの object schema で表現してみるか、property を raw passthrough（最小実装）で済ませるかの 2 択。openpyxl の `tests/data/genuine` に custom.xml を持つフィクスチャは無いため、自前で簡易 fixture を仕込む。続けて §7 utils（coordinate / datetime / units / inference / escape）。
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
- [~] §6 packaging 層（manifest + relationships + docProps/core.xml + docProps/app.xml 完了：Schema に `raw` ElementDef 種別追加（XmlNode subtree を opaque 保持、HeadingPairs/TitlesOfParts/HLinks/Hyperlinks/DigSig 等の vt:vector 系を実装ぬきで round-trip）、`ExtendedProperties` 28 フィールド、DEFAULT_PREFIXES の `XPROPS_NS` / `CUSTPROPS_NS` を `''`（default namespace）に修正。openpyxl `genuine/empty.xlsx` の `docProps/app.xml` で application/docSecurity/scaleCrop/company/appVersion 等の値一致 + headingPairs/titlesOfParts の raw 保持 (vt:vector size 一致) + 完全 round-trip。169 tests pass。残：docProps/custom.xml）
- [ ] §7 utils（coordinate, datetime, units, inference, escape, exceptions）
- [ ] §8 compat
- [ ] §9 phase-1 テスト群
- [ ] §10 フェーズ1 完了条件

### フェーズ2 以降は到達時に展開する。

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
