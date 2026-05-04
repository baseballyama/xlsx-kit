# PROGRESS

`/loop` 自走モードでフェーズ1→7 を順に実装するための、ターン横断の状態ファイル。
**唯一の正は `docs/plan/`**。本ファイルは「いまどこまで終わったか」だけを記録する。

## カレント

- **フェーズ**: フェーズ1（基盤層）
- **次のタスク**: フェーズ1 §3 XML 層の続き：iterparse SAX。`src/xml/iterparse.ts` で `iterParse(stream | bytes): AsyncIterableIterator<SaxEvent>` を `saxes` ベースで実装。SaxEvent は `{ kind: 'start' | 'end' | 'text', name?, attrs?, text? }` でよい。namespace 解決は parser と同等の挙動（Clark）。phase 4 の read-only worksheet で消費される。受け入れ条件は 1k 行 sheetData の全 cell カウントが一致するスモークテスト。saxes が dynamic import で済むなら `openxml-js` のメインエントリには載せず streaming サブパス側に隔離。
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
- [~] §3 XML 層（namespaces / tree / parser DOM / serializer DOM 完了：`serializeXml(node, opts?): Uint8Array`、Clark 表記から prefix 復元（DEFAULT_PREFIXES → 未登録 NS は `ns{N}`）、`xml` prefix は予約・宣言不要、root NS が DEFAULT で `''` の時のみ default として emit、attr は never default、`& < > " ' \\r \\n \\t` のエスケープ、XML 宣言は opt 切替可。`parseXml → serializeXml → parseXml` で `xl/workbook.xml` / `[Content_Types].xml` / `_rels/.rels` の round-trip pass。105 tests pass。残：iterparse SAX / canonical compare helper / 大規模 round-trip）
- [ ] §4 Schema 層（Schema 型 + `toTree`/`fromTree`）
- [ ] §5 XmlStreamWriter
- [ ] §6 packaging 層（manifest, relationships, doc properties）
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
