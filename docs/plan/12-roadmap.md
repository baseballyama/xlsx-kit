# 12. ロードマップ / マイルストーン / リスク

## 1. フェーズと見積もり

> 単位は "人週"。1 名フルタイム前提。複数人並行で短縮可能（並行可な箇所は注記）。

| フェーズ | 主題 | 期間 | バージョン目安 | 並行可 |
|---------|------|------|--------------|-------|
| フェーズ1 | 基盤層（IO / ZIP / XML / Schema / packaging / utils） | 4〜5週 | `0.1.0-alpha` | ZIP と XML と Schema は別人で同時可 |
| フェーズ2 | コアモデル（Cell / Style / Stylesheet / Worksheet データ / Formula） | 4〜5週 | `0.2.0-alpha` | スタイルと数式は別 |
| フェーズ3 | read / write 通常モード | 5〜6週 | `0.3.0-alpha` | reader と writer は別人可（ただし Schema 完成が前提） |
| フェーズ4 | streaming（read-only / write-only） | 3週 | `0.4.0-alpha` | – |
| フェーズ5 | リッチ機能（コメント / リンク / テーブル / 検証 / 条件付き / 定義名 / 印刷） | 5週 | `0.5.0-beta` | サブシステム単位で並列可 |
| フェーズ6 | **チャート全機能 + 描画** | **8〜12週** | `0.7.0-beta` | chart 種類ごとに並列可 |
| フェーズ7 | ピボット / VBA / OLE passthrough | 2〜3週 | `0.9.0-rc` | – |
| 仕上げ | ドキュメント / 移行ガイド / QA / パフォーマンスチューニング | 2〜3週 | `1.0.0` | – |

合計 **33〜42 人週**。1名 単独で 8〜10ヶ月、2名並行で 5〜6ヶ月、3名で 4〜5ヶ月の見積もり。

## 2. 各フェーズのゲート

| ゲート | 通過条件 |
|-------|---------|
| **G1: 基盤完成** | フェーズ1 完了。ZIP/XML/Schema/manifest が round-trip 成立 |
| **G2: モデル完成** | フェーズ2 完了。JSON ラウンドトリップ可、stylesheet dedup pass |
| **G3: 最小 read/write** | フェーズ3 完了。openpyxl の `genuine/sample.xlsx` round-trip pass |
| **G4: 大規模対応** | フェーズ4 完了。100 万 cell の read/write が 1GB 以下のヒープで pass |
| **G5: リッチ機能完備** | フェーズ5 完了。テーブル / 条件付き書式 / 定義名 が round-trip + 編集 API 動作 |
| **G6: チャート全機能** | フェーズ6 完了。25 chart 種類 round-trip + 編集、視覚 QA pass |
| **G7: 互換性最終確認** | フェーズ7 完了。VBA / pivot / customUI が壊れない |
| **G8: 1.0 リリース** | typedoc / 移行ガイド / size budget / CI 全通過 |

## 3. リスク管理

### 3.1 高リスク

| リスク | 影響 | 緩和策 |
|--------|------|-------|
| ChartML の物量がフェーズ6 を膨れさせる | 全体スケジュール遅延 | フェーズ6 を「フル対応」と「拡張 chartex」に分離。前者を `1.0`、後者を `1.1` に切り出す選択肢を残す |
| ECMA-376 と Excel 365 の挙動差 | 出力 xlsx が Excel で開けない / 警告 | フェーズ6 で QA 自動化（Excel・LibreOffice・Google Sheets を回す） |
| 数式 tokenizer の精度不足 | shared formula を読むと壊れる | フェーズ2 で openpyxl のテストフィクスチャを完全移植して pass を必須化 |
| ブラウザ環境の差異（Web Streams 実装ブレ） | Safari でストリームが動かない | フェーズ4 でフォールバック（Buffer 経由）を必ず実装 |
| 性能リグレッション気付かず | 後段で取り返せない | フェーズ1 から bench を CI に乗せ、25% 劣化で fail |

### 3.2 中リスク

| リスク | 影響 | 緩和策 |
|--------|------|-------|
| openpyxl の参考実装が ECMA-376 と微妙に違う | 互換性のグレーゾーン | 差分を発見したら ECMA-376 を採用、テストで明文化 |
| サブパス export の bundler 互換性問題（Vite / webpack 5 / esbuild / parcel） | 利用者の DX 悪化 | フェーズ3 で各 bundler のスモークテストを CI 化 |
| TypeScript エラーの大量発生 | 型修正で時間消費 | strict から始める。中途で緩めない |
| pnpm / Node のバージョン差 | local と CI のズレ | volta / asdf を docs に明記 |

### 3.3 低リスク

| リスク | 影響 | 緩和策 |
|--------|------|-------|
| openpyxl 側の更新（3.1.x → 3.2.x） | submodule の固定方針が必要 | 特定 tag 固定。差分が必要時のみ手動更新 |
| 外部依存（fflate 等）の breaking | 上流追随コスト | major 更新は別 PR で慎重に |

## 4. 完了の定義（Definition of Done）

### 4.1 1.0 リリースの条件

- [ ] フェーズ1〜7 すべて完了
- [ ] 公開 API のドキュメント整備（typedoc + ガイド）
- [ ] 移行ガイド（openpyxl → xlsxify）の整備
- [ ] バンドルサイズ予算遵守（[01-architecture.md](./01-architecture.md) §13、[11-build-publish.md](./11-build-publish.md) §1.4）
- [ ] CI 全通過（typecheck / lint / test / browser test / size / bench）
- [ ] openpyxl の `genuine/`, `reader/`, `writer/`, `chart/`, `drawing/` 系フィクスチャ全件で round-trip pass
- [ ] 視覚 QA で Excel / LibreOffice / Google Sheets / WPS で chart が等価表示
- [ ] エラー階層が docs に明記
- [ ] THIRD_PARTY_NOTICES.md に openpyxl 含む全依存の著作権表示

### 4.2 1.0 後の継続項目（v1.x）

- パフォーマンス改善：worksheet の行/列に対する `Float64Array` SoA 化（数百万セル超え向け）
- xsd ベースのバリデータ（オプトイン）
- `dataframe-to-rows` 相当の interop ヘルパ（`@stdlib/...` / arquero / nodejs-polars 連携）
- カスタム XML パーツの構造化編集
- 暗号化された xlsx の **読み込み**（`msoffcrypto-tool` 互換 JS 実装が出てきたら検討）
- ピボットテーブル構造編集 API
- データモデル / Power Pivot 拡張のサポート

## 5. 進行フロー

1. **キックオフ**: 本ドキュメントを read-through。フェーズ1 のチェックリストを ToDo 化。
2. **週次**: フェーズ内のチェックリストを進捗チェック。完了タスクは PR と紐付け。
3. **フェーズゲート**: 上記 G1〜G8 をクリアしたら次へ。各ゲートで以下を実施：
   - リグレッションテスト全件
   - サイズ予算チェック
   - サンプルアプリで動作確認
   - フィードバックを issue 化
4. **リリース**: changeset で発行。GitHub Release ノートに「openpyxl 何相当の機能をサポート」を明記。

## 6. 並列開発のための分割案

複数人で進める場合の分担例：

### 2 名チームの場合（5〜6ヶ月）
- **Dev A**: フェーズ1（基盤）→ フェーズ2 のスタイル → フェーズ3 reader → フェーズ5 → フェーズ6 のチャート schema
- **Dev B**: フェーズ1 の XML/Schema → フェーズ2 の Workbook/Worksheet/Formula → フェーズ3 writer → フェーズ4 streaming → フェーズ6 の drawing/anchor/cx → フェーズ7

### 3 名チームの場合（4〜5ヶ月）
- **Dev A**（Core）: 基盤 + コアモデル + reader/writer
- **Dev B**（Streaming + Rich）: streaming + コメント + テーブル + 条件付き書式 + 定義名
- **Dev C**（Visual）: チャート全種 + drawing + chartex + chartsheet

並列は **G2 通過後** から本格化できる。それまでは Schema を共同で作るため逐次。

## 7. 具体的な「明日からの一歩」

**Day 1**:
1. submodule を `git submodule update --init --recursive` で展開済みであることを確認
2. `pnpm init` でリポジトリ初期化
3. `tsconfig.json` / `biome.json` / `vitest.config.ts` / `tsup.config.ts` を [11-build-publish.md](./11-build-publish.md) の通りに作成
4. `package.json` に [11-build-publish.md](./11-build-publish.md) §1.1 の骨子をコピー
5. `src/` のスケルトンディレクトリを作成（[01-architecture.md](./01-architecture.md) §3）
6. `THIRD_PARTY_NOTICES.md` を作成し、openpyxl の MIT ライセンスを明記
7. CI（GitHub Actions）の最小設定を `.github/workflows/ci.yml` に置く
8. GitHub repo の branch protection を main に設定（typecheck/lint/test 必須）

**Day 2 以降**:
- フェーズ1 の §1（IO 抽象）から着手
- 各タスクは GitHub issue 化、フェーズ番号 + 目印 label を付ける（`phase-1`, `phase-2`, …）
- すべての変更は PR、最低 1 reviewer、squash merge
