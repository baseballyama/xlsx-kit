# 04. フェーズ2: コアモデル（Cell / Style / Stylesheet）

**目的**: ワークブック内部で完結するデータモデルを実装する。読み込み・書き出しの XML 経路は次フェーズ。
**期間目安**: 4〜5週間
**前提**: フェーズ1 が完了していること
**完了条件**: 「メモリ上で Workbook を組み立て、StyleSheet を整合させ、JSON として round-trip できる」状態。

## 1. 全体方針

- すべて plain object + 純粋関数（[01-architecture.md](./01-architecture.md) §5）。
- イミュータブル要素（Font/Fill/Border/Side/Alignment/Protection/NumberFormat/Color）は freeze されたオブジェクト。`make*` で作って以後変更しない。
- ミュータブル要素（Cell/Worksheet/Workbook）は属性アクセスで mutate する慣習を保つ。
- Stylesheet は **値ベースの dedup** を ID 振りで行う（同値判定のキャッシュキー = JSON 化された normalized 値）。

## 2. Cell モデル（`src/cell/`）

### 2.1 型

```ts
export type CellDataType = 'n' | 's' | 'b' | 'd' | 'inlineStr' | 'str' | 'e' | 'formula';
//                     number  shared  bool  date  inline-str  str   error  formula

export type CellValue =
  | number
  | string
  | boolean
  | Date
  | { kind: 'duration'; ms: number }
  | { kind: 'error'; code: ExcelErrorCode }
  | { kind: 'rich-text'; runs: TextRun[] }
  | { kind: 'formula'; formula: string; cachedValue?: number | string | boolean; t?: 'array' | 'shared' | 'normal' | 'dataTable'; ref?: string; si?: number }
  | null;

export interface Cell {
  /** 1-based row */
  row: number;
  /** 1-based column */
  col: number;
  /** 値そのもの。型と合わせて保存（dataType を別個に持たない、value から都度導出） */
  value: CellValue;
  /** Stylesheet.cellXfs index。0 = 既定 */
  styleId: number;
  /** ハイパーリンク (worksheet 側でも管理されるが、cell が知っていると便利) */
  hyperlinkId?: number;
  /** legacy comment（threaded comment は worksheet 側に持たせる） */
  comment?: number; // commentId への参照
}

export type ExcelErrorCode =
  | '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!'
  | '#NAME?' | '#NUM!'   | '#N/A'    | '#GETTING_DATA';

export interface MergedCell extends Cell { merged: true; }
```

### 2.2 関数群

```ts
// src/cell/cell.ts
export function makeCell(row: number, col: number, value?: CellValue, styleId = 0): Cell;
export function getCoordinate(c: Cell): string;             // "A1"
export function setCellValue(c: Cell, v: CellValue): void;  // 必要なエスケープ・推論を行わない（後述 2.3）
```

### 2.3 値の取り扱い方針

- `setCellValue` は **型推論を行わない**。呼び出し側が明示的に変換する。理由：
  - openpyxl の `_bind_value` 相当を毎回走らせるのはホットパスで重い。
  - TS では型システムでカバーできる。
- 高レベル API として **別名** で推論ありの helper を用意：
  ```ts
  // src/cell/value.ts
  export function bindValue(c: Cell, v: number | string | boolean | Date): void;
  ```
- `bindValue` は string が `=...` で始まれば formula、`#...` の error code、`Date` なら date、それ以外は number/string/bool。
- 数式・配列数式・データテーブル数式は **明示的な API**：
  ```ts
  export function setFormula(c: Cell, formula: string, opts?: { cachedValue?: number | string | boolean }): void;
  export function setArrayFormula(c: Cell, ref: string, formula: string, opts?: { cachedValue?: number | string | boolean }): void;
  export function setSharedFormula(c: Cell, si: number, formula?: string, ref?: string, opts?: { cachedValue?: number | string | boolean }): void;
  ```

### 2.4 Rich text

```ts
// src/cell/rich-text.ts
export interface InlineFont {
  name?: string; sz?: number; b?: boolean; i?: boolean; u?: 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  strike?: boolean; vertAlign?: 'baseline' | 'superscript' | 'subscript';
  color?: Color;
  /* … */
}
export interface TextRun { text: string; font?: InlineFont; }
export type RichText = ReadonlyArray<TextRun>;
```

`Cell.value` に `{ kind: 'rich-text', runs }` で格納。

### 2.5 受け入れ条件

- [ ] `Cell` の生成・読み出し・更新で `value` の型情報が壊れない（unit test）
- [ ] `bindValue` の挙動が openpyxl の `_bind_value` と等価（datetime, formula, error 各ケース）
- [ ] `getCoordinate(makeCell(1,1)) === 'A1'`、`getCoordinate(makeCell(1,16384)) === 'XFD1'`
- [ ] `MergedCell` 用の `value === null` を厳密に強制

## 3. Style モデル（`src/styles/`）

### 3.1 値オブジェクト一覧

| 型 | フィールド要点 |
|----|---------------|
| `Color` | `rgb?: string` (`AARRGGBB`) / `theme?: number` / `indexed?: number` / `auto?: boolean` / `tint?: number` |
| `Font` | name, size, bold, italic, underline, strike, color, scheme, family, charset, vertAlign |
| `Side` | style (`thin`/`medium`/.../`double`), color |
| `Border` | left, right, top, bottom, diagonal, vertical, horizontal, diagonalUp, diagonalDown, outline |
| `PatternFill` | patternType, fgColor, bgColor |
| `GradientFill` | type, stops (degree/left/right/top/bottom)、stops 配列 |
| `Fill` = `PatternFill | GradientFill` |
| `Alignment` | horizontal, vertical, indent, textRotation, wrapText, shrinkToFit, justifyLastLine, readingOrder |
| `Protection` | locked, hidden |
| `NumberFormat` | `numFmtId: number; formatCode: string` |

すべて **readonly + freeze**。spread で更新。

```ts
// src/styles/fonts.ts
export interface Font {
  readonly name?: string;
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  readonly strike?: boolean;
  readonly color?: Color;
  readonly family?: number;
  readonly charset?: number;
  readonly scheme?: 'major' | 'minor';
  readonly vertAlign?: 'baseline' | 'superscript' | 'subscript';
}
export function makeFont(opts?: Partial<Font>): Font {
  return Object.freeze({ ...opts });
}
export const DEFAULT_FONT: Font = makeFont({ name: 'Calibri', size: 11, family: 2, scheme: 'minor' });
```

### 3.2 Color

theme カラーは **解決しない**。書き戻し時に `theme` index をそのまま吐く。tint は浮動小数点の round-trip で誤差が出やすいので **文字列保存**（小数点 16 桁）にしておく実装も可。

`hex` の rgb は openpyxl と同じく **8 桁 (alpha 含む)** を canonical 形にする。`'FFFF0000'` 等。

### 3.3 NumberFormat

- `BUILTIN_FORMATS`（id 0〜163）は `src/styles/numbers.ts` に const として持つ。openpyxl の `BUILTIN_FORMATS` を完全コピーする。
- カスタム書式 id は **164 以上**を自動採番。
- `isDateFormat(code)`：openpyxl と同じ正規表現ヒューリスティクスを移植（`is_date_format`）。
- `isTimedeltaFormat(code)`：同じく `is_timedelta_format`。

### 3.4 Stylesheet 内部構造

```ts
// src/styles/stylesheet.ts
export interface Stylesheet {
  fonts: Font[];                    // index 0 が default
  fills: Fill[];                    // index 0,1 が none/gray125 既定
  borders: Border[];                // index 0 が default
  numFmts: Map<number, string>;     // id → formatCode
  cellXfs: CellXf[];                // 各 cell が参照する style index 集合
  cellStyleXfs: CellXf[];           // 名前付きスタイルのテンプレート
  namedStyles: NamedStyle[];        // name + cellXf の組
  dxfs: DifferentialStyle[];        // 条件付き書式・テーブル用
  tableStyles: TableStyle[];        // 表スタイル
  colors?: ColorList;               // indexed パレット override
  /** 値→indexの dedup マップ (in-memory のみ。XML 化されない) */
  _fontIdByKey: Map<string, number>;
  _fillIdByKey: Map<string, number>;
  _borderIdByKey: Map<string, number>;
  _xfIdByKey: Map<string, number>;
}

export interface CellXf {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  xfId?: number;                  // cellStyleXfs を参照する index
  alignment?: Alignment;
  protection?: Protection;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyNumberFormat?: boolean;
  applyAlignment?: boolean;
  applyProtection?: boolean;
  pivotButton?: boolean;
  quotePrefix?: boolean;
}
```

### 3.5 Stylesheet 操作 API

```ts
// すべて純粋関数だが、引数 ss を mutate する（性能上）
export function makeStylesheet(): Stylesheet;
export function addFont(ss: Stylesheet, f: Font): number;
export function addFill(ss: Stylesheet, f: Fill): number;
export function addBorder(ss: Stylesheet, b: Border): number;
export function addNumFmt(ss: Stylesheet, formatCode: string): number;   // 既定形式は dedup
export function addCellXf(ss: Stylesheet, xf: CellXf): number;
export function addCellStyleXf(ss: Stylesheet, xf: CellXf): number;
export function addNamedStyle(ss: Stylesheet, name: string, xf: CellXf, builtinId?: number): number;
```

dedup キー戦略：
- `Font` / `Fill` / `Border` / `Alignment` / `Protection` は **JSON.stringify** で正規化（フィールドを sort）してキー化。`stableStringify` を utils に置く。
- `CellXf` は (fontId, fillId, borderId, numFmtId, xfId, JSON of alignment, JSON of protection, applyFlags) を結合した string をキーに。

### 3.6 Cell ↔ Stylesheet ブリッジ（`src/styles/cell-style.ts`）

```ts
export function getCellFont(wb: Workbook, c: Cell): Font;
export function getCellFill(wb: Workbook, c: Cell): Fill;
export function getCellBorder(wb: Workbook, c: Cell): Border;
export function getCellAlignment(wb: Workbook, c: Cell): Alignment;
export function getCellProtection(wb: Workbook, c: Cell): Protection;
export function getCellNumberFormat(wb: Workbook, c: Cell): string;

export function setCellFont(wb: Workbook, c: Cell, font: Font): void;
export function setCellFill(wb: Workbook, c: Cell, fill: Fill): void;
export function setCellBorder(wb: Workbook, c: Cell, border: Border): void;
export function setCellAlignment(wb: Workbook, c: Cell, alignment: Alignment): void;
export function setCellProtection(wb: Workbook, c: Cell, protection: Protection): void;
export function setCellNumberFormat(wb: Workbook, c: Cell, formatCode: string): void;
```

実装：
1. プールから現在の `cellXfs[c.styleId]` を取得。
2. 該当フィールドだけ差し替えた新 `CellXf` を作る。
3. `addCellXf` で dedup → 新 index。
4. `c.styleId` を更新。

`setRange*` ヘルパは `src/worksheet/style-range.ts` で worksheet 側 import に置く（worksheet を import しない使い方ではバンドル増えない）。

### 3.7 Built-in styles (`src/styles/builtins.ts`)

openpyxl `styles/builtins.py` の各 NamedStyle を **オブジェクトリテラル** として定義：

```ts
export const BUILTIN_NAMED_STYLES = {
  Normal:    { font: makeFont({ name: 'Calibri', size: 11 }), /* … */ },
  'Good':    { /* … */ },
  'Bad':     { /* … */ },
  'Neutral': { /* … */ },
  'Calculation': { /* … */ },
  'Check Cell':  { /* … */ },
  'Currency':    { /* … */ },
  'Currency [0]':{ /* … */ },
  'Comma':       { /* … */ },
  'Comma [0]':   { /* … */ },
  'Percent':     { /* … */ },
  'Linked Cell': { /* … */ },
  'Hyperlink':   { /* … */ },
  'Followed Hyperlink': { /* … */ },
  'Note':        { /* … */ },
  'Title':       { /* … */ },
  'Headline 1':  { /* … */ },
  'Headline 2':  { /* … */ },
  'Headline 3':  { /* … */ },
  'Headline 4':  { /* … */ },
  /* Accent1〜6, Total, Warning Text, Input, Output, Explanatory Text, etc. */
} as const;
```

利用者が require した時に Stylesheet にレジスタする helper：

```ts
export function ensureBuiltinStyle(wb: Workbook, name: keyof typeof BUILTIN_NAMED_STYLES): number;
```

### 3.8 Differential Style (`src/styles/differential.ts`)

`DifferentialStyle` は `Font` / `Fill` / `Border` / `Alignment` / `Protection` / `NumberFormat` のうち**指定されたフィールドのみ**を持つ partial style。条件付き書式・テーブルが参照する。

```ts
export interface DifferentialStyle {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  protection?: Protection;
  numFmt?: { id?: number; code: string };
}

export function addDxf(ss: Stylesheet, dxf: DifferentialStyle): number;
```

### 3.9 受け入れ条件

- [ ] BUILTIN_FORMATS の全 id が openpyxl と一致
- [ ] `addFont(ss, makeFont({ bold: true }))` を 1000 回呼んでも `ss.fonts.length === 2`（default + bold）
- [ ] CellXf の dedup が openpyxl 出力と等価（同等の Workbook を構築すると同じ cellXfs プールになる）
- [ ] Color の rgb/theme/indexed/auto/tint の変換が round-trip 一致

## 4. Workbook / Worksheet データモデル（`src/workbook/`, `src/worksheet/`）

### 4.1 Workbook 型

```ts
// src/workbook/workbook.ts
export interface Workbook {
  /** Sheet の表示順 */
  sheets: SheetRef[];
  activeSheetIndex: number;
  /** workbook プロパティ */
  properties: CoreProperties;
  customProperties?: CustomProperties;
  appProperties?: ExtendedProperties;
  workbookPr: { date1904: boolean; codeName?: string; /* … */ };
  calcPr: CalcProperties;
  bookViews: BookView[];
  definedNames: DefinedName[];
  externalLinks: ExternalLink[];
  /** Stylesheet（プール込） */
  styles: Stylesheet;
  /** sharedStrings の order を保持（書き戻し時に必要） */
  sharedStrings: { table: string[]; index: Map<string, number> };
  /** thread comment authors */
  authors: string[];
  /** 元 zip を keep_vba 相当で保持（passthrough 用） */
  vbaProject?: Uint8Array;
  vbaSignature?: Uint8Array;
  /** theme XML を保持（編集しないなら原文で round-trip） */
  themeXml?: Uint8Array;
  /** 未対応要素の passthrough バケット */
  passthrough: Map<string, Uint8Array>;
  /** logger */
  _logger: Logger;
}

export type SheetRef =
  | { kind: 'worksheet'; sheet: Worksheet; sheetId: number; rId: string; state: 'visible' | 'hidden' | 'veryHidden' }
  | { kind: 'chartsheet'; sheet: Chartsheet; sheetId: number; rId: string; state: 'visible' | 'hidden' | 'veryHidden' };
```

### 4.2 Workbook free function

```ts
export function createWorkbook(opts?: { date1904?: boolean }): Workbook;
export function addWorksheet(wb: Workbook, title: string, opts?: { index?: number }): Worksheet;
export function addChartsheet(wb: Workbook, title: string): Chartsheet;
export function removeSheet(wb: Workbook, title: string): void;
export function getSheet(wb: Workbook, title: string): Worksheet | Chartsheet | undefined;
export function setActiveSheet(wb: Workbook, title: string): void;
export function copyWorksheet(wb: Workbook, sourceTitle: string, newTitle?: string): Worksheet;
export function defineName(wb: Workbook, name: string, formula: string, opts?: { scope?: string }): void;
```

### 4.3 Worksheet 型

```ts
// src/worksheet/worksheet.ts
export interface Worksheet {
  title: string;
  /** Cell ストア。row→col→Cell の二段 Map */
  rows: Map<number, Map<number, Cell>>;
  /** 行/列のメタ */
  columnDimensions: Map<number, ColumnDimension>;
  rowDimensions: Map<number, RowDimension>;
  defaultColumnWidth?: number;
  defaultRowHeight?: number;
  /** マージ */
  mergedCells: MultiCellRange;
  /** ビュー */
  views: SheetView[];
  /** ハイパーリンク */
  hyperlinks: Hyperlink[];
  /** データ検証 */
  dataValidations: DataValidation[];
  /** 条件付き書式 */
  conditionalFormatting: ConditionalFormattingList;
  /** Excel テーブル */
  tables: TableDefinition[];
  /** AutoFilter */
  autoFilter?: AutoFilter;
  /** 印刷 */
  pageSetup: PageSetup;
  pageMargins: PageMargins;
  printOptions: PrintOptions;
  headerFooter: HeaderFooter;
  rowBreaks: PageBreak[];
  colBreaks: PageBreak[];
  /** 図形・画像・チャート */
  drawings: Drawing[];
  /** legacy comments / threaded comments */
  legacyComments: LegacyComment[];
  threadedComments: ThreadedComment[];
  /** 印刷タイトル / 印刷範囲 (定義名へ転送) */
  printArea?: string;
  printTitleRows?: string;
  printTitleCols?: string;
  /** 凍結ペイン */
  freezePanes?: string;
  /** 保護 */
  sheetProtection?: SheetProtection;
  /** 親 workbook back ref（mutate 用; 循環参照 OK：ハンドル限定） */
  workbook: Workbook;
}
```

`workbook` の back-ref は **JSON.stringify を妨げる** ので、シリアライズが必要なら `replacer` で除外する。

### 4.4 Worksheet 関数群

```ts
// セル取得・更新
export function getCell(ws: Worksheet, row: number, col: number): Cell | undefined;
export function setCell(ws: Worksheet, row: number, col: number, value: CellValue, styleId?: number): Cell;
export function deleteCell(ws: Worksheet, row: number, col: number): void;
export function appendRow(ws: Worksheet, values: CellValue[]): number; // 新規 row index を返す
export function iterRows(ws: Worksheet, opts?: { minRow?: number; maxRow?: number; minCol?: number; maxCol?: number; valuesOnly?: boolean }): IterableIterator<Cell[]>;

// マージ
export function mergeCells(ws: Worksheet, ref: string): void;
export function unmergeCells(ws: Worksheet, ref: string): void;

// 列幅・行高
export function setColumnWidth(ws: Worksheet, col: number, width: number): void;
export function setRowHeight(ws: Worksheet, row: number, height: number): void;
export function hideColumn(ws: Worksheet, col: number): void;

// 凍結ペイン
export function setFreezePanes(ws: Worksheet, ref: string): void;

// 印刷
export function setPrintArea(ws: Worksheet, ref: string): void;
```

### 4.5 cell-range / multi-cell-range

```ts
// src/worksheet/cell-range.ts
export interface CellRange { minRow: number; minCol: number; maxRow: number; maxCol: number; }
export function parseRange(s: string): CellRange;        // "A1:B5"
export function rangeToString(r: CellRange): string;     // "A1:B5"
export function rangeContainsCell(r: CellRange, row: number, col: number): boolean;
export function rangeContainsRange(a: CellRange, b: CellRange): boolean;
export function shiftRange(r: CellRange, dr: number, dc: number): CellRange;
export function unionRange(a: CellRange, b: CellRange): CellRange;
export function intersectionRange(a: CellRange, b: CellRange): CellRange | null;
export function* iterRangeCoordinates(r: CellRange): IterableIterator<{ row: number; col: number }>;
```

`MultiCellRange` は `CellRange[]` の小ラッパ。

### 4.6 受け入れ条件

- [ ] `appendRow` 100 万回のスループットが M1 上で > 200k rows/s
- [ ] `iterRows({ valuesOnly: true })` の出力順が row → col の昇順
- [ ] mergeCells を行うと該当範囲の左上以外の Cell が `MergedCell` として再生成される（openpyxl の `MergedCellRange.format()` 等価）
- [ ] freeze pane の "B2" などの設定が SheetView の pane オブジェクトに正しく反映される

## 5. 数式モデル（`src/formula/`）

### 5.1 Tokenizer

openpyxl `formula/tokenizer.py` を移植。**正規表現 + state machine** の素直な実装。

```ts
// src/formula/tokenizer.ts
export type TokenType =
  | 'OPERAND' | 'OPERATOR' | 'FUNC' | 'PAREN' | 'SEP' | 'WSPACE' | 'OPEN-ARRAY' | 'CLOSE-ARRAY';
export type TokenSubtype =
  | 'RANGE' | 'NUMBER' | 'STRING' | 'BOOL' | 'ERROR' | 'OPEN' | 'CLOSE' | 'PREFIX' | 'INFIX' | 'POSTFIX' | 'ARG';

export interface Token { value: string; type: TokenType; subtype?: TokenSubtype; }

export function tokenize(formula: string): Token[];
```

### 5.2 Translator

shared formula を per-cell に展開する。openpyxl `formula/translate.py:102-134` を移植。

```ts
// src/formula/translator.ts
export function translateFormula(formula: string, originRC: { row: number; col: number }, destRC: { row: number; col: number }): string;
```

### 5.3 受け入れ条件

- [ ] openpyxl の `formula/tests/data/*.txt` フィクスチャ全件を移植して tokenizer 結果が一致
- [ ] shared formula の translator が openpyxl と一致
- [ ] 数式評価は **絶対にしない**（errata: 評価器を入れる気配があれば PR を blocker レビュー）

## 6. JSON ラウンドトリップ

フェーズ2の完了確認として、以下が動くこと：

```ts
const wb = createWorkbook();
const ws = addWorksheet(wb, 'Sheet1');
setCell(ws, 1, 1, 42);
setCellFont(wb, getCell(ws, 1, 1)!, makeFont({ bold: true }));
mergeCells(ws, 'A2:B3');

const json = JSON.stringify(wb, jsonReplacer); // 循環参照を除去するヘルパ
const wb2 = JSON.parse(json, jsonReviver) as Workbook;
expect(getCell(wb2.sheets[0].sheet as Worksheet, 1, 1)?.value).toBe(42);
```

これにより、データ層単独でテスト・デバッグできる状態になる。

## 7. テスト（`tests/phase-2/`）

- `tests/phase-2/cell.test.ts`
- `tests/phase-2/styles/*.test.ts`（font, fill, border, alignment, protection, color, numbers, builtins）
- `tests/phase-2/stylesheet-dedup.test.ts`
- `tests/phase-2/worksheet.test.ts`
- `tests/phase-2/cell-range.test.ts`
- `tests/phase-2/formula-tokenize.test.ts`
- `tests/phase-2/formula-translate.test.ts`
- `tests/phase-2/json-roundtrip.test.ts`

property-based: `tests/phase-2/dedup.property.ts` で fast-check により dedup の冪等性を確認。

## 8. 完了条件（フェーズ2 全体）

- [ ] §2〜§6 各受け入れ条件
- [ ] フェーズ1 のテストが回帰していない
- [ ] バンドルサイズ予算違反なし（[01-architecture.md](./01-architecture.md) §13）
- [ ] サンプルスクリプト: 「createWorkbook → 100 セル書き込み → JSON シリアライズ → 復元 → 同一性検証」が end-to-end pass
