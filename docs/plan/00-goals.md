# 00. ゴール / 非ゴール / 互換性方針

## 1. プロダクトのゴール

- **`openpyxl` の主要機能を TypeScript で再実装し、Node 22+ と主要モダンブラウザの両方で動かす。**
- API 形状は openpyxl の Python API になるべく忠実に従う。ただし、Pythonic な慣習（snake_case、kwargs、descriptor、`Worksheet[A1]` のような糖衣）は **JS/TS 慣習に置き換える**：
  - 命名: `camelCase`（クラス名は `PascalCase`）
  - 引数: 末尾オプションオブジェクト
  - インデキシング: `ws.cell(1, 1)` のメソッド形式優先（`ws['A1']` も `Proxy` で提供）
  - 戻り値: 例外ベース。`Result` 型は採用しない（openpyxl 風）
- **既存 xlsx ファイルの read → 編集 → write のラウンドトリップで、未編集部分のバイト/XML 同等性を可能な限り保つ**。

## 2. プロダクトの非ゴール

| 機能 | 採否 | 備考 |
|------|------|------|
| 数式の評価（calc engine） | ❌ | openpyxl も非対応。値は cached value をそのまま使う。 |
| `.xls` (BIFF) サポート | ❌ | openpyxl も非対応。 |
| `.csv` 直接サポート | △ | 取り込み/書き出しユーティリティとして将来検討 |
| 暗号化された xlsx の復号 | ❌ | パスワード保護されたシート構造の保持は対応 |
| 編集中ファイルのリアルタイム共有 | ❌ | 範囲外 |
| Excel 独自の formatted text rendering | ❌ | データ層のみ。レンダリングは扱わない |
| Cython 級のネイティブ高速化 | △ | 必要なら WebAssembly (sax-wasm) の利用を検討 |

## 3. 想定ユースケース（優先度順）

1. **小〜中規模 xlsx のサーバサイド生成**（API レスポンス、レポート出力）
2. **クライアントサイドでの xlsx エクスポート**（ブラウザで Blob を生成）
3. **既存 xlsx の編集・追記**（テンプレ運用）
4. **大規模 xlsx の読み取り** (read-only / streaming) — 数百万セル規模
5. **大規模 xlsx の書き出し** (write-only / streaming) — 数百万セル規模
6. **チャート・画像を含む xlsx の round-trip 編集**

## 4. 互換性ターゲット

### 4.1 ランタイム

| ランタイム | バージョン | 備考 |
|-----------|-----------|------|
| Node.js | **18 LTS 以降** | `node:fs/promises`, `Web Streams API`, `globalThis.crypto` 必須 |
| Deno | 1.40+ (ベストエフォート) | ESM のみ。`node:` インポート互換層に依存 |
| Bun | 1.0+ (ベストエフォート) | Node 互換 API を使用 |
| Chrome / Edge | 最新 - 2 | ESM, top-level await, Web Streams |
| Firefox | 最新 - 2 | 同上 |
| Safari | 16+ | 同上 |
| iOS Safari | 16+ | 同上 |

> 古い Node や IE などはサポート対象外。トランスパイル設定でカバーしない。

### 4.2 言語/ツール

| 項目 | バージョン |
|------|----------|
| TypeScript | **5.4 以上** |
| `target` | `ES2022`（top-level await、`Error.cause`、`#private` 利用前提） |
| `module` | `ESNext`（バンドル時に解決） |
| `strict` | `true` |
| ESM 優先 | あり（`type: module`）|
| CJS 互換 | あり（dual export） |

### 4.3 OOXML 仕様

- **対象**: ECMA-376 Edition 1〜4、Strict と Transitional の両方を read 対応、write は **Transitional 既定**（openpyxl と同じ）。
- 拡張機能（threadedComments、外部参照など）は段階的にサポート。

## 5. openpyxl 機能の優先度

| カテゴリ | Phase | サポートレベル | openpyxl 参照 |
|---------|-------|---------------|--------------|
| ZIP/Manifest/Relationships | 1 | フル | `packaging/`, `xml/` |
| Cell value (str/num/bool/date/formula) | 2 | フル | `cell/` |
| Style (font/fill/border/alignment/numFmt/protection) | 2 | フル | `styles/` |
| NamedStyle / 既定スタイル | 2 | フル | `styles/named_styles.py`, `styles/builtins.py` |
| Worksheet (rows/cols/merge/freeze/hide) | 3 | フル | `worksheet/worksheet.py` |
| 数式（trans 含む shared/array） | 3 | フル（評価なし） | `formula/` |
| Workbook 読み書き | 3 | フル | `reader/excel.py`, `writer/excel.py` |
| Read-only / Write-only モード | 4 | フル | `worksheet/_read_only.py`, `_write_only.py` |
| コメント (legacy + threaded) | 5 | フル | `comments/` |
| ハイパーリンク | 5 | フル | `worksheet/hyperlink.py` |
| テーブル / オートフィルタ | 5 | フル | `worksheet/table.py`, `filters.py` |
| データ検証 | 5 | フル | `worksheet/datavalidation.py` |
| 条件付き書式 | 5 | フル | `formatting/` |
| 定義名 (Defined Names) | 5 | フル | `workbook/defined_name.py` |
| 外部リンク | 5 | パススルー round-trip | `workbook/external_link/` |
| 印刷設定 / ヘッダフッタ | 5 | フル | `worksheet/page.py`, `header_footer.py` |
| ページ区切り | 5 | フル | `worksheet/pagebreak.py` |
| 画像 (PNG/JPEG/GIF/BMP/WebP/TIFF/SVG) | 6 | フル | `drawing/image.py` |
| Drawing (Picture/Shape/GroupShape/Connector) | 6 | フル | `drawing/spreadsheet_drawing.py` |
| **チャート全機能（全種類 + 全装飾要素）** | 6 | **フル（最重要）** | `chart/` |
| チャート: Bar / Line / Pie / Scatter / Area | 6 | フル | – |
| チャート: Radar / Bubble / Doughnut / Stock / Surface | 6 | フル | – |
| チャート: 3D 全 variants（barChart3D / lineChart3D / pie3DChart / area3DChart / surfaceChart / surface3DChart） | 6 | フル | `chart/_3d.py` |
| チャート: 軸（Category/Value/Date/Series）+ 副軸 | 6 | フル | `chart/axis.py` |
| チャート: Trendline / ErrorBars / UpDownBars / DropLines / HiLowLines / Marker | 6 | フル | `chart/trendline.py` ほか |
| チャート: DataLabel / DataPoint / 個別書式 / 凡例位置・カスタム位置 | 6 | フル | `chart/label.py`, `legend.py` |
| チャート: ShapeProperties（fill/line/effect/3D bevel/glow/shadow） | 6 | フル | `chart/shapes.py`, `drawing/effect.py` |
| Chartsheet | 6 | フル | `chartsheet/` |
| ピボットテーブル | 7 | パススルー round-trip | `pivot/` |
| VBA macro 保存 (`keep_vba`) | 7 | パススルー（バイナリ温存） | `reader/excel.py:162-165` |
| ActiveX / OLE / ribbon customUI | 7 | パススルー | `reader/excel.py:101-104` |
| ドキュメントプロパティ (core/app/custom) | 3 | フル | `packaging/core.py` ほか |

「パススルー round-trip」= XML/バイナリは保持して再書き出しできるが、TS 側で構造を編集する API は提供しない（あるいは限定的）。

## 6. 互換性の度合い

- **read fidelity**: openpyxl が読める xlsx は原則すべて読めるべき。読めない場合はテストフィクスチャを openpyxl 側でも追加し、共通の挙動とする。
- **write fidelity**: openpyxl 出力との byte-identical は **目標ではない**（openpyxl 自身も lxml/ElementTree 切り替えで揺れる）。代わりに次を保証する：
  1. 出力 xlsx は **Excel 365 / LibreOffice / Google Sheets** で問題なく開ける
  2. **canonical XML compare**（属性順序・空白を正規化）で同型である
  3. round-trip 後にフィクスチャの **意味的内容** が保たれる
- **API surface compat with openpyxl**: 1:1 の API ミラーは目指さない。ただし主要クラス（`Workbook`, `Worksheet`, `Cell`, `Font`, `Border`, ...）と主要メソッド（`load_workbook`/`save`、`worksheet.cell`、`worksheet.append`）は同名で提供する。

## 7. 参照管理

- `reference/openpyxl/` を **読み取り専用** として運用。コミットしない。
- 移植中は openpyxl 側の `file:line` を **必ず TS 実装のコメントに残す**。例:
  ```ts
  // Port of openpyxl/descriptors/base.py:28-50 (Typed validator).
  ```
- openpyxl の MIT ライセンスは継承する。`LICENSE` には現在の TS 著者表記があるが、`THIRD_PARTY_NOTICES.md` を別途追加して openpyxl の著作権表示を含める（フェーズ 1 完了時）。
