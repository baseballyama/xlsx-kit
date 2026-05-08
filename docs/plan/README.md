# xlsxify 実装計画

`openpyxl`（Python, MIT, 3.1.5 系）を TypeScript に移植し、**Node 18+ / 主要モダンブラウザ** の両方で動作する OOXML (xlsx/xlsm) 操作ライブラリ `xlsxify` を作るための、フェーズ分割された実装計画ドキュメントです。

このディレクトリは「これに従えば実装が進む」レベルの設計仕様として整備しています。フェーズごとの**成果物・受け入れ条件**を明記しているため、進捗状況の確認にも使えます。

## 文書構成

| # | ファイル | 内容 |
|---|----------|------|
| 00 | [00-goals.md](./00-goals.md) | プロジェクトのゴール、非ゴール、対象ユースケース、互換性方針 |
| 01 | [01-architecture.md](./01-architecture.md) | 全体アーキテクチャ、パッケージ構成、Node/ブラウザ両対応戦略、外部依存選定 |
| 02 | [02-mapping.md](./02-mapping.md) | openpyxl Python モジュール ↔ xlsxify TS モジュールの対応表 |
| 03 | [03-foundations.md](./03-foundations.md) | フェーズ1: 基盤層（ZIP、XML、I/O 抽象、Schema/Descriptor 相当） |
| 04 | [04-core-model.md](./04-core-model.md) | フェーズ2: 値モデル（Cell, Coordinate, Range, Style, Stylesheet） |
| 05 | [05-read-write.md](./05-read-write.md) | フェーズ3: workbook / worksheet の read / write 実装 |
| 06 | [06-streaming.md](./06-streaming.md) | フェーズ4: 大規模ファイル向けストリーミング（read-only / write-only） |
| 07 | [07-rich-features.md](./07-rich-features.md) | フェーズ5: コメント、ハイパーリンク、テーブル、データバリデーション、条件付き書式、定義名、外部リンク |
| 08 | [08-charts-drawings.md](./08-charts-drawings.md) | フェーズ6: 画像・描画・チャート（DrawingML / ChartML） |
| 09 | [09-pivot-vba.md](./09-pivot-vba.md) | フェーズ7: ピボットテーブル / VBA / 暗号化（パススルー） |
| 10 | [10-testing.md](./10-testing.md) | テスト戦略（単体・golden・ファジング・ブラウザ・パフォーマンス） |
| 11 | [11-build-publish.md](./11-build-publish.md) | ビルド・公開（dual ESM/CJS、Conditional Exports、サブパス、CI/CD） |
| 12 | [12-roadmap.md](./12-roadmap.md) | ロードマップ、マイルストーン、見積もり、リスク管理 |

## 用語ノート

- **openpyxl**: Python 製の OOXML ライブラリ。本リポの `reference/openpyxl/` 配下に submodule で取り込み済み（quintagroup/openpyxl ミラー、3.1.5 系）。
- **OOXML / SpreadsheetML**: ECMA-376 で標準化された Office Open XML のスプレッドシート部分。
- **Workbook / Worksheet / Chartsheet / Cell / Style**: OOXML の概念。openpyxl のクラス名と原則一致させる。
- **Descriptor**: openpyxl 独自の宣言的 XML マッピング機構（`openpyxl/descriptors/`）。本ポートでは「Schema」と呼ぶ TS 実装で置き換える。
- **Phase N**: 段階的開発のフェーズ番号。後段は前段の成果に依存する。

## 進め方

1. **必ず順番にフェーズを進める** — 後段は前段の API/型に依存している。
2. **各フェーズの「受け入れ条件」を満たしたら次へ** — 中途半端な進行はテストで検出されるよう、phase ごとに `tests/phase-N/` を作成する。
3. **openpyxl 側の挙動に疑問があれば必ず `reference/openpyxl/` の該当ファイル:行 を読む** — このリポジトリのドキュメントは要約に過ぎない。
4. **未対応機能は明示的に `NotImplementedError` 相当を投げる** — silent failure にしない。
