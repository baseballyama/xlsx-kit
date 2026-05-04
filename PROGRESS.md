# PROGRESS

`/loop` 自走モードでフェーズ1→7 を順に実装するための、ターン横断の状態ファイル。
**唯一の正は `docs/plan/`**。本ファイルは「いまどこまで終わったか」だけを記録する。

## カレント

- **フェーズ**: 0 (bootstrap) → フェーズ1 へ
- **次のタスク**: フェーズ1 §1 (IO 抽象 `XlsxSource` / `XlsxSink`) の最小実装
- **ブランチ**: `main`（直接 commit 運用、squash 不要）

## 完了履歴

各エントリは「フェーズ §X.Y タスク → コミット hash」の形式。

### bootstrap

- [x] **bootstrap**: 計画コミット（`docs:` / 4573 行） — `773ae69`
- [x] **bootstrap**: TS プロジェクト雛形（`package.json` / `tsconfig.json` / `tsconfig.build.json` / `biome.json` / `.npmrc` / `.nvmrc` / `src/index.ts` placeholder / `THIRD_PARTY_NOTICES.md` / typescript 5.9 + @types/node 22 install / `pnpm typecheck` pass）
- [x] **bootstrap**: vitest 設定 + smoke test（vitest 4.1 + @vitest/coverage-v8、`tests/phase-0/smoke.test.ts` で `pnpm test` pass、`pnpm typecheck` pass）
- [ ] **bootstrap**: tsup 設定（`tsup.config.ts` / `pnpm build` smoke）
- [ ] **bootstrap**: biome lint 通過（既存ファイルが pass する状態）
- [ ] **bootstrap**: GitHub Actions CI 雛形（typecheck / lint / test ジョブ）

### フェーズ1: 基盤層（[03-foundations.md](docs/plan/03-foundations.md)）

- [ ] §1 I/O 抽象（`XlsxSource` / `XlsxSink`、Node ヘルパ、ブラウザヘルパ）
- [ ] §2 ZIP 層（fflate ベース reader/writer、ストリーミング）
- [ ] §3 XML 層（fast-xml-parser DOM + saxes SAX + namespace 定数）
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
