# 01. アーキテクチャ

## 1. レイヤード設計

```
┌────────────────────────────────────────────────────────────────┐
│  Public API: Workbook / Worksheet / Cell / Font / Border / ... │
├────────────────────────────────────────────────────────────────┤
│  Domain layer: 値モデル + StyleProxy + StyleArray              │
│  (cell, styles, worksheet, chart, drawing, comments, formula)  │
├────────────────────────────────────────────────────────────────┤
│  Serialization layer: Schema (descriptor 相当) + XmlSerializer │
│  (各クラスの toTree / fromTree、ネームスペース、name 解決)      │
├────────────────────────────────────────────────────────────────┤
│  XML layer: parse / serialize / iterparse                      │
│  (fast-xml-parser + saxes ベースの薄いラッパ)                   │
├────────────────────────────────────────────────────────────────┤
│  Package layer: ZIP archive + manifest + relationships         │
│  (fflate ベース。ストリームは Web Streams API)                  │
├────────────────────────────────────────────────────────────────┤
│  IO layer: Source / Sink 抽象（File / Buffer / Blob / Stream）  │
└────────────────────────────────────────────────────────────────┘
```

各レイヤは **下層への単方向依存** とする。テストはレイヤごとに分けて書く（`tests/io`, `tests/xml`, `tests/serialization`, `tests/domain`, `tests/public`）。

## 2. Node / ブラウザ両対応戦略

### 2.1 環境差分の隔離

`io/` レイヤだけが環境を意識する。それ以外のレイヤはランタイムに依存しない。

| 機能 | Node 実装 | ブラウザ実装 |
|------|----------|--------------|
| バイナリ I/O 入力 | `Buffer`, `Uint8Array`, `node:fs` (path/file handle), `Readable` | `Uint8Array`, `Blob/File`, `ReadableStream` |
| バイナリ I/O 出力 | `Buffer`, `Uint8Array`, `node:fs`, `Writable` | `Uint8Array`, `Blob`, `WritableStream` |
| ハッシュ（SHA1, MD5 等） | `node:crypto` | `globalThis.crypto.subtle` |
| 乱数 | `node:crypto` | `globalThis.crypto` |

### 2.2 dual entry pattern

`package.json` の `exports` フィールドで `node` / `browser` 条件を切り替える：

```jsonc
{
  "name": "ooxml-js",
  "type": "module",
  "exports": {
    ".": {
      "types": "./dist/index.d.ts",
      "node": {
        "import": "./dist/index.node.mjs",
        "require": "./dist/index.node.cjs"
      },
      "browser": "./dist/index.browser.mjs",
      "default": "./dist/index.browser.mjs"
    },
    "./streaming": {
      "types": "./dist/streaming.d.ts",
      "node": {
        "import": "./dist/streaming.node.mjs",
        "require": "./dist/streaming.node.cjs"
      },
      "browser": "./dist/streaming.browser.mjs"
    }
  },
  "sideEffects": false
}
```

`io/` 配下に `io/node.ts` / `io/browser.ts` を置き、コア層からは `io/index.ts` 経由で参照する。バンドラがエントリ条件で適切に解決する。

### 2.3 Web Streams API を共通言語に

Node 18+ は `ReadableStream`/`WritableStream` をネイティブで提供する。ooxml-js の **ストリーミング API は Web Streams を共通インターフェース** として採用し、Node 専用の `Readable`/`Writable` は `Readable.toWeb()` で橋渡しする。

```ts
// 共通シグネチャ
interface XlsxSource {
  toBytes(): Promise<Uint8Array>;
  toStream(): ReadableStream<Uint8Array>;
}
```

## 3. パッケージレイアウト（モノレポは採用しない）

> **方針**: 公開パッケージは `ooxml-js` の **単一パッケージ**。内部ディレクトリで分離し、サブパス公開で粒度を出す。複数パッケージにすると submodule との対応が崩れて移植コストが跳ねるため、当面は単一パッケージで進める。

```
ooxml-js/
├── reference/openpyxl/       # submodule（参照のみ）
├── docs/plan/                # 本計画ドキュメント
├── docs/api/                 # 自動生成（typedoc）
├── src/
│   ├── index.ts              # 公開エントリ
│   ├── io/
│   │   ├── source.ts         # XlsxSource interface
│   │   ├── sink.ts           # XlsxSink interface
│   │   ├── node.ts           # Node 専用実装
│   │   └── browser.ts        # ブラウザ専用実装
│   ├── zip/
│   │   ├── reader.ts         # 解凍
│   │   ├── writer.ts         # 圧縮
│   │   └── streaming.ts      # ストリーム書き出し
│   ├── xml/
│   │   ├── parser.ts         # parseXml / fromString
│   │   ├── serializer.ts     # toString / writer
│   │   ├── iterparse.ts      # SAX-style iterator
│   │   ├── namespaces.ts     # 全 OOXML namespace 定数
│   │   └── tree.ts           # 内部 XML node 表現（軽量 DOM 風）
│   ├── schema/
│   │   ├── core.ts           # Schema, Field, Validator
│   │   ├── descriptors.ts    # Typed/Bool/Integer/Float/String/Set/...
│   │   ├── nested.ts         # Nested* 相当
│   │   ├── sequence.ts       # Sequence/NestedSequence/MultiSequence
│   │   ├── serialisable.ts   # Serializable mixin (toTree/fromTree)
│   │   └── alias.ts          # Alias 相当
│   ├── packaging/
│   │   ├── manifest.ts       # [Content_Types].xml
│   │   ├── relationships.ts  # *.rels
│   │   ├── core.ts           # docProps/core.xml
│   │   ├── extended.ts       # docProps/app.xml
│   │   └── custom.ts         # docProps/custom.xml
│   ├── compat/
│   │   ├── numbers.ts        # Decimal & numeric helpers
│   │   └── singleton.ts
│   ├── utils/
│   │   ├── coordinate.ts     # A1↔(row,col)
│   │   ├── datetime.ts       # Excel serial date
│   │   ├── units.ts          # EMU / pt / px
│   │   ├── inference.ts      # value type inference
│   │   ├── escape.ts         # XML char escaping
│   │   ├── exceptions.ts     # 例外クラス
│   │   └── indexed-list.ts
│   ├── cell/
│   │   ├── cell.ts           # Cell + MergedCell
│   │   ├── read-only.ts      # ReadOnlyCell
│   │   ├── rich-text.ts      # CellRichText / TextBlock
│   │   └── writer.ts         # cell → XML 出力
│   ├── styles/
│   │   ├── fonts.ts
│   │   ├── fills.ts
│   │   ├── borders.ts
│   │   ├── alignment.ts
│   │   ├── protection.ts
│   │   ├── numbers.ts
│   │   ├── colors.ts
│   │   ├── named-styles.ts
│   │   ├── builtins.ts
│   │   ├── differential.ts
│   │   ├── style-array.ts
│   │   ├── style-proxy.ts
│   │   ├── stylesheet.ts
│   │   └── table.ts
│   ├── formatting/
│   │   ├── conditional.ts
│   │   └── rule.ts
│   ├── workbook/
│   │   ├── workbook.ts
│   │   ├── child.ts
│   │   ├── defined-name.ts
│   │   ├── properties.ts
│   │   ├── protection.ts
│   │   ├── views.ts
│   │   ├── external-link/
│   │   ├── reader.ts
│   │   └── writer.ts
│   ├── worksheet/
│   │   ├── worksheet.ts
│   │   ├── read-only.ts
│   │   ├── write-only.ts
│   │   ├── reader.ts
│   │   ├── writer.ts
│   │   ├── dimensions.ts
│   │   ├── cell-range.ts
│   │   ├── multi-cell-range.ts
│   │   ├── merge.ts
│   │   ├── filters.ts
│   │   ├── data-validation.ts
│   │   ├── hyperlink.ts
│   │   ├── header-footer.ts
│   │   ├── page-break.ts
│   │   ├── page.ts
│   │   ├── pivot.ts
│   │   ├── properties.ts
│   │   ├── protection.ts
│   │   ├── related.ts
│   │   ├── table.ts
│   │   ├── views.ts
│   │   ├── controls.ts
│   │   └── ole.ts
│   ├── comments/
│   │   ├── comment.ts
│   │   ├── comment-sheet.ts
│   │   └── author.ts
│   ├── formula/
│   │   ├── tokenizer.ts
│   │   ├── translator.ts
│   │   └── tokens.ts
│   ├── drawing/
│   │   ├── spreadsheet-drawing.ts
│   │   ├── image.ts
│   │   ├── picture.ts
│   │   ├── shape.ts
│   │   ├── geometry.ts
│   │   ├── colors.ts
│   │   ├── fill.ts
│   │   ├── line.ts
│   │   ├── effect.ts
│   │   ├── text.ts
│   │   ├── anchor.ts
│   │   ├── xdr.ts
│   │   └── relation.ts
│   ├── chart/
│   │   ├── chart-space.ts
│   │   ├── chart-base.ts
│   │   ├── bar-chart.ts
│   │   ├── line-chart.ts
│   │   ├── pie-chart.ts
│   │   ├── scatter-chart.ts
│   │   ├── area-chart.ts
│   │   ├── radar-chart.ts
│   │   ├── bubble-chart.ts
│   │   ├── doughnut-chart.ts
│   │   ├── stock-chart.ts          # passthrough
│   │   ├── surface-chart.ts        # passthrough
│   │   ├── 3d.ts
│   │   ├── axis.ts
│   │   ├── series.ts
│   │   ├── reference.ts
│   │   ├── data-source.ts
│   │   ├── plot-area.ts
│   │   ├── legend.ts
│   │   ├── title.ts
│   │   ├── layout.ts
│   │   ├── marker.ts
│   │   ├── label.ts
│   │   ├── trendline.ts
│   │   ├── error-bar.ts
│   │   └── reader.ts
│   ├── chartsheet/
│   │   ├── chartsheet.ts
│   │   ├── views.ts
│   │   ├── properties.ts
│   │   └── protection.ts
│   ├── pivot/
│   │   ├── cache.ts          # passthrough
│   │   ├── records.ts        # passthrough
│   │   ├── table.ts
│   │   ├── fields.ts
│   │   └── common.ts
│   ├── streaming/
│   │   ├── read-only-workbook.ts
│   │   ├── write-only-workbook.ts
│   │   └── xml-stream-writer.ts
│   └── public/
│       ├── load.ts           # loadWorkbook(source, opts): Promise<Workbook>
│       └── save.ts           # saveWorkbook(wb, sink, opts): Promise<void>
├── tests/                    # vitest 単体テスト（src 構造とミラー）
├── tests/fixtures/           # openpyxl データを参照 + 追加フィクスチャ
├── tests/golden/             # canonical XML / SHA256 ハッシュ
├── tests/browser/            # Playwright + vitest browser
├── tests/perf/               # ベンチマーク（mitata or tinybench）
├── scripts/                  # ビルド・リリース・フィクスチャ生成
├── package.json
├── tsconfig.json
├── tsconfig.build.json
├── vitest.config.ts
├── vitest.workspace.ts
├── .github/workflows/        # CI
├── biome.json                # Biome を採用（lint+format 一本化）
└── README.md
```

> ファイル単位で 1ファイル ≤ ~400 行の方針。openpyxl の Python 1ファイルが大きい場合、TS 側で機能単位に分割してよい（その分割は本ドキュメントの範囲外）。

## 4. 外部依存の選定

### 4.1 ランタイム依存（最小化方針）

| openpyxl の依存 | TS 代替 | 採用 | 備考 |
|----------------|--------|------|------|
| `et_xmlfile`（streaming write） | 自作 `XmlStreamWriter`（後述） | ✅ | 数百行で済むため自作が最適 |
| `lxml`（高速 parse） | 不要 | — | TS では fast-xml-parser で十分 |
| `Pillow`（画像） | `image-size` | ✅ | PNG/JPEG/GIF/BMP/WebP のヘッダ寸法取得 |
| `defusedxml`（XXE 対策） | XML パーサ側で `processEntities: false` 等 | ✅ | パーサ設定だけで済むため別 lib 不要 |
| `numpy`（数値型） | 不要 | — | JS は number/BigInt のみ。Decimal 必要なら decimal.js |
| `pandas`（DataFrame） | 不要 | — | `dataframe_to_rows` 相当はオプション |

### 4.2 採用ライブラリ（決定）

| 領域 | ライブラリ | バージョン | 採用理由 |
|------|----------|-----------|----------|
| ZIP 圧縮/解凍 | **fflate** | ^0.8 | pure JS、ESM、tree-shake 可、Web Streams 互換、~30KB |
| XML parse（DOM） | **fast-xml-parser** | ^4.4 | pure JS、tree-shake 可、Node/ブラウザ両対応 |
| XML parse（SAX/streaming） | **saxes** | ^6 | pure JS、Web Streams にラップ可、TypeScript 型あり |
| 画像メタ情報 | **image-size** | ^1.1 | PNG/JPEG/GIF/BMP/WebP/TIFF/SVG。pure JS |
| Decimal（必要時） | **decimal.js-light** | ^2 | ピボットや precise number で必要時のみ動的 import |

> **alternative library 採用に変更があれば、本ドキュメントの 4.2 と `package.json` を同時に書き換えること。**

### 4.3 開発依存

| 領域 | ツール | 備考 |
|------|------|------|
| パッケージマネージャ | **pnpm** | workspaces 不要だが速度・lockfile 厳密性で採用 |
| テストランナー | **vitest** | Node + jsdom + browser mode + bench。ESM ネイティブ |
| ブラウザテスト | **@vitest/browser + Playwright** | Chromium/Firefox/WebKit 全部回す |
| プロパティテスト | **fast-check** | descriptor の round-trip ファジング |
| Lint/Format | **Biome** | ESLint + Prettier の代替。高速。設定一本化 |
| 型チェック | **tsc --noEmit** | CI で実行 |
| ビルド | **tsup** | esbuild ベース。dual ESM/CJS、d.ts 生成、外部条件付きエントリ |
| バンドルサイズ計測 | **size-limit** | CI で size budget を強制 |
| API ドキュメント | **typedoc** | `docs/api/` 出力 |
| Changelog | **changeset** | リリースフロー自動化 |

## 5. データ指向設計（クラス禁止 / 純粋関数主体）

ツリーシェイキングと実行性能を最優先するため、**`class` キーワードは原則使わない**。状態は **plain object + interface 型**、操作は **副作用のない free function** で表現する。これにより以下が成り立つ：

- **未使用関数はバンドラが除去できる**（クラスメソッドは「クラスを参照した時点で全メソッド込み」になりがちで除去されにくい）。
- **シェイプが安定する**（V8 の hidden class 最適化が効きやすい）。
- **継承・this バインドの実行コストが消える**。
- **シリアライズ／デシリアライズが楽**（`structuredClone` / JSON / WASM 境界を越える時に問題が起きない）。

### 5.1 ルール

1. **`class` を書かない**。例外は標準 `Error` 派生（`OpenXmlError extends Error`）のみ。
2. **データは plain object**。`interface` または `type` で形を定める。
   ```ts
   export interface Font {
     readonly name?: string;
     readonly size?: number;
     readonly bold?: boolean;
     readonly italic?: boolean;
     readonly color?: Color;
     // …
   }
   ```
3. **値オブジェクトは凍結（`Object.freeze`）または `readonly` 型** にする。
4. **構築は `make*` 関数**で行う。例：`makeFont(opts: Partial<Font>): Font`。デフォルト値の埋め込みと freeze はここで行う。
5. **更新は immutable spread**：`const f2 = { ...f1, bold: true }`。`with` ヘルパは作らない（呼び出しコストとサイズの無駄）。
6. **Workbook 直下のミューテーションだけは性能上 mutable** にする（数百万セルの書き換えで spread すると致命的）。`Workbook` `Worksheet` `Cell` は内部的に可変。ただしスタイル・フォント等のプール参照は immutable。
7. **インスタンスメソッドは持たない**。代わりにモジュールスコープで関数として export：
   ```ts
   // src/cell/cell.ts
   export interface Cell { row: number; col: number; value: CellValue; styleId: number; /* … */ }
   export function makeCell(row: number, col: number, value?: CellValue): Cell;
   export function setCellValue(cell: Cell, value: CellValue): void;
   export function getCoordinate(cell: Cell): string;
   ```
8. **「`cell.font = newFont`」スタイルは提供しない**。Workbook のスタイルプールに対する index 操作が必要なため、必ず free function 経由：
   ```ts
   import { setCellFont } from 'ooxml-js/cell-style';
   setCellFont(workbook, cell, { bold: true, color: { rgb: 'FF0000' } });
   ```
9. **「`workbook.save()`」スタイルは提供しない**。`saveWorkbook(wb, sink)` のみ。
10. **継承は使わない**。共通フィールドは交差型（`A & B`）かフラット展開で表現。継承の代わりに **discriminated union**：
    ```ts
    type Anchor =
      | { kind: 'absolute'; pos: Point2D; ext: PositiveSize2D; /* … */ }
      | { kind: 'oneCell';  from: Marker; ext: PositiveSize2D; /* … */ }
      | { kind: 'twoCell';  from: Marker; to: Marker; /* … */ };
    ```

### 5.2 ergonomic-first な公開 API

ユーザに `setCellValue(cell, 42)` を強要すると Excel 風のシンプルさが失われる。**最も多用されるアクセスは plain getter/setter** のまま受け入れる：

```ts
// OK: 純粋な data field の代入は性能・サイズに無害
cell.value = 42;
worksheet.title = 'Summary';

// NG: スタイルプールを伴う操作はメソッド化しない、関数化する
setCellFont(wb, cell, { bold: true });
addImage(ws, image, 'A1');
```

つまり「値の書き換えは property assignment、構造を伴う操作は free function」という二段構えとする。

### 5.3 Schema (descriptor 相当) も純データ

XML マッピング層は **クラスのフィールドではなく、別ファイルの `const` Schema** として定義する。Cell や Font 自身は schema を知らない。

```ts
// src/styles/borders.ts
export interface Border { /* … */ }
export function makeBorder(opts?: Partial<Border>): Border;

// src/styles/borders.schema.ts  ← serialize 時のみ import される
import type { Border } from './borders.js';
export const BorderSchema: Schema<Border> = {
  tagname: 'border',
  attrs: {
    diagonalUp:   { kind: 'bool',  optional: true },
    diagonalDown: { kind: 'bool',  optional: true },
    outline:      { kind: 'bool',  default: true },
  },
  elements: [
    { name: 'left',       kind: 'object', schema: () => SideSchema, optional: true },
    { name: 'right',      kind: 'object', schema: () => SideSchema, optional: true },
    { name: 'top',        kind: 'object', schema: () => SideSchema, optional: true },
    { name: 'bottom',     kind: 'object', schema: () => SideSchema, optional: true },
    { name: 'diagonal',   kind: 'object', schema: () => SideSchema, optional: true },
    { name: 'vertical',   kind: 'object', schema: () => SideSchema, optional: true },
    { name: 'horizontal', kind: 'object', schema: () => SideSchema, optional: true },
  ],
};
```

Schema を別ファイル（`*.schema.ts`）に切ることで、**「値だけ使うが書き出さない」コードパスでは Schema をバンドルから完全に除去できる**（ESM の sideEffect-free + named import で達成）。

シリアライザは単一の純粋関数：
```ts
export function toTree<T>(value: T, schema: Schema<T>): XmlNode;
export function fromTree<T>(node: XmlNode, schema: Schema<T>): T;
```

`toTree`/`fromTree` の実装も内部で `if/else` 連鎖（switch on `kind`）にし、validator メソッドの仮想呼び出しを避ける。

### 5.4 StyleArray / スタイル参照

Cell はスタイル本体を持たない。**Stylesheet プール上の index** を 1 個の数値 (`xfId`) または 9 整数のタプルで保持する。openpyxl は 9 整数の `array.array` を使うが、TS では **数値 1 個（`xfId`）+ プール側で展開** に簡略化する：

```ts
export interface Cell {
  row: number;
  col: number;
  value: CellValue;
  /** index into workbook._styles.cellXfs; 0 = default */
  styleId: number;
  /** optional inline style override (rare path) */
  inlineStyleId?: number;
  // …
}
```

書き込み API：
```ts
// Workbook 上の全 Cell に対して一括スタイル変更
setCellFont(wb, cell, font);  // wb._styles.fonts に dedup add → cellXfs に dedup add → cell.styleId 更新
setRangeFont(ws, 'A1:C10', font);
```

これにより、**スタイル変更を伴わないフローでは `setCellFont` を一切 import しない** ＝ そのコードパスは tree-shake で消える。

## 6. ツリーシェイキング戦略

### 6.1 module 分割粒度

サブパス export を粒度細かく切り、利用者は必要な部分だけ import できるようにする。`package.json` の `exports` に明示する。

```jsonc
"exports": {
  ".":               "./dist/index.js",            // Workbook + Worksheet + Cell までのコアバンドル
  "./read":          "./dist/read.js",             // load only
  "./write":         "./dist/write.js",            // save only
  "./streaming":     "./dist/streaming.js",        // read-only / write-only
  "./styles":        "./dist/styles.js",           // Font / Fill / Border / Alignment / NumberFormat
  "./conditional":   "./dist/conditional.js",      // 条件付き書式
  "./formula":       "./dist/formula.js",          // tokenizer / translator
  "./drawing":       "./dist/drawing.js",          // 画像・図形
  "./chart":         "./dist/chart.js",            // チャート
  "./pivot":         "./dist/pivot.js",            // ピボット passthrough
  "./schema":        "./dist/schema.js"            // Schema / toTree / fromTree（lower-level）
}
```

利用者：
```ts
// 値読み出しだけしたい人（最小バンドル）
import { loadWorkbook, getCell } from 'ooxml-js/read';

// 書き込みフルセット
import { createWorkbook, addSheet, save } from 'ooxml-js';
import { setCellFont, defaultBorder } from 'ooxml-js/styles';
```

### 6.2 sideEffect 設定

- `package.json` に `"sideEffects": false` を必ず入れる。
- 例外的に side-effect を持つファイルがあれば配列で列挙する。
- `Object.freeze` は副作用に見えるが ESM の static import 解析では問題にならない。

### 6.3 dynamic import で「重いモジュール」を切り離す

以下は dynamic import を採用する：
- **Chart 系の writer**（chart 種別ごとに schema が大きい）：`save` 中に `await import('ooxml-js/chart-writer')`
- **画像エンコード（image-size）**：drawing が無いワークブックには load されない
- **Pivot writer / VBA passthrough**：該当機能が使われた時のみ
- **数式 tokenizer / translator**：shared formula を含むセルを読んだ時だけ

ストリーミング読み出しでは worksheet ごとに動的 import で必要な reader だけロードする。

### 6.4 export pattern

- すべて **named export** のみ。`export default` は使わない（ESM で tree-shake が効きにくいケースがある）。
- barrel file（`index.ts` で再 export）を作る場合、**そこに副作用のあるコードを入れない**。barrel 自体は `"sideEffects": false` 前提でないと意味がない。

## 7. パフォーマンス指針

### 7.1 ホットパス（数百万セル走査）

- `for (let i = 0; i < arr.length; i++)` の素朴ループを優先。`for...of` / `forEach` は cold path のみ。
- 配列・オブジェクトの shape を **生成時から固定**。途中で `delete` や新フィールド追加をしない（V8 megamorphic 化を避ける）。
- セルは **連続する `Float64Array` 系の SoA**（structure-of-arrays）に格納する場合があるが、デフォルトは `Map<string, Cell>` 相当の **行ごとの Map<col, Cell>**。
- 文字列値は **shared strings インデックス** を保持して dedup（openpyxl と同じ）。
- 数式 token / Cell value union 判定は **discriminated union + switch** で書く。`instanceof` は使わない。

### 7.2 XML write のホットパス

- 文字列連結ではなく **chunk 配列 + `Uint8Array` への TextEncoder 直書き**。
- worksheet の cell 1 つあたりの XML 出力は **template 関数 1 回** で完結させる（switch 1 回、分岐最小化）。
- 共通文字列（`<row r="`, `"><c r="`, `" t="s"><v>`）は **module-level の const 文字列 / バイト列** にして再利用。
- `XMLSerializer` 系の DOM ベース API は使わない（速度が劣る）。

### 7.3 XML read のホットパス

- worksheet（`xl/worksheets/sheet*.xml`）は **SAX で iterparse**（saxes）。フル DOM はメモリ的にも速度的にも不利。
- styles.xml / workbook.xml / [Content_Types].xml など小さい XML は **fast-xml-parser で一括 parse**。
- shared strings は SAX 経由で **`Uint32Array` index** に流し込む。

### 7.4 数値・日付

- **`Date` を使わない方が速いケースが多い**。Excel serial（`number`）のまま保持し、ユーザが要求した時だけ `excelSerialToDate(serial)` を呼ぶ。
- 1900 leap-bug 補正は分岐 1 個。早期 return。
- 整数だけが期待される箇所では `| 0` で int32 化を強制（V8 の Smi 最適化）。

### 7.5 ZIP ストリーム

- fflate の **streaming API（`Zip`/`AsyncZip`/`Unzip`）** を使う。`Uint8Array` 全展開モードは大規模ファイルで OOM する。
- ZIP entry のメタを **書き出し開始時点で確定** させ、deflate と並行してセルを生成する。

## 8. StyleProxy ではなく cellStyle*Functions*

openpyxl の `StyleProxy` は `cell.font` のオブジェクトアクセス糖衣を提供するが、TS では非対応とする：

| 操作 | API |
|------|-----|
| 取得 | `getCellFont(wb, cell): Font` |
| 設定 | `setCellFont(wb, cell, font): void` |
| 範囲一括 | `setRangeFont(ws, ref, font): void` |

それぞれ独立した named export → tree-shake 可能。ユーザが `setCellFont` を使わなければ、Stylesheet pool の add ロジックもバンドルから消える。

## 9. エラー設計

- 公開 API は **`OpenXmlError` 階層** を投げる（`OpenXmlIoError`, `OpenXmlSchemaError`, `OpenXmlInvalidWorkbookError`, `OpenXmlNotImplementedError`, ...）。
- 内部例外は `cause` でチェーン。Node 18+ / モダンブラウザは `Error.cause` をサポート。
- `NotImplementedError` 相当は **明示的に投げる**。silent fallback はしない。

## 10. ロギング

- ロガーは **dependency injection**。標準は `console.warn`。`createWorkbook({ logger: { warn, info, debug } })` で差し替え可。
- 既定で「未対応の OOXML 拡張要素」「strict round-trip できない要素」は `warn` を出す。

## 11. メモリ戦略

- 通常 mode（フルロード）: 全セル・全スタイルを in-memory 保持。
- read-only mode: ZIP は `XlsxSource` のままに保持し、worksheet は `iterRows()` 呼び出し時に SAX で再パース。（[06-streaming.md](./06-streaming.md) §2）
- write-only mode: `appendRow` / `appendCell` は `XmlStreamWriter` 経由で **直接 ZIP のエントリにストリーム書き込み**。openpyxl は temp ファイルを使うが、TS では Web Streams で繋ぐ。（[06-streaming.md](./06-streaming.md) §3）

## 12. CI / 品質ゲート

PR ごとに以下を必須：
1. `pnpm typecheck` — 型エラー 0
2. `pnpm lint` — Biome エラー 0
3. `pnpm test` — Node + jsdom 環境で全テスト pass
4. `pnpm test:browser` — Chromium/Firefox/WebKit でコアテスト pass
5. `pnpm size` — バンドルサイズ予算オーバーなし
6. `pnpm bench:smoke` — perf ベンチ regression なし（base からの 25% 劣化を fail とする）

## 13. バンドルサイズ予算（指針）

| エントリ | サイズ予算（min+gz） |
|---------|---------------------|
| `ooxml-js`（フル） | ≤ 200 KB |
| `ooxml-js/streaming`（read/write only サブセット） | ≤ 80 KB |
| `ooxml-js/light`（read-only + 値読み出しのみ・将来検討） | ≤ 50 KB |
