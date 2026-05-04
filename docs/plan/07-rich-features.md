# 07. フェーズ5: リッチ機能（コメント / ハイパーリンク / テーブル / 検証 / 条件付き書式 / 定義名）

**目的**: フェーズ3 で passthrough 扱いにした非中核機能を、TS で完全に **構造化** する。
**期間目安**: 5週間
**前提**: フェーズ1〜4
**完了条件**: 各機能が「読み込み → 編集 API → 書き戻し」全てで動く。

機能ごとに独立に進められる。複数人で並行作業する場合のメリットあり。

---

## 1. コメント

### 1.1 二系統

| 種類 | OOXML 部品 | TS 構造 |
|------|-----------|--------|
| Legacy comment | `xl/commentsN.xml` + `xl/drawings/vmlDrawingN.vml` | `LegacyComment` |
| Threaded comment | `xl/threadedCommentsN.xml` + `xl/persons/personN.xml` | `ThreadedComment` |

### 1.2 LegacyComment

```ts
// src/comments/comment.ts
export interface LegacyComment {
  ref: string;                // "A1"
  authorIndex: number;        // workbook.authors[]
  text: RichText | string;
  /** 表示用 widht/height (pt) */
  width?: number;
  height?: number;
  /** VML shape の anchor。round-trip のため保存 */
  vmlShapeXml: string;
}
```

VML 部分は **structure を完全には解析しない**。round-trip のため XML フラグメントを保持。書き戻し時に `vmlShapeId` を更新する程度。

### 1.3 ThreadedComment

```ts
// src/comments/threaded-comment.ts
export interface ThreadedComment {
  id: string;                  // GUID
  ref: string;
  personId: string;
  text: string;
  createdOn: Date;
  parentId?: string;           // reply の場合
  done?: boolean;
  mentions?: Array<{ mentionId: string; personId: string; range: { start: number; len: number } }>;
}

export interface Person {
  id: string;
  displayName: string;
  userId: string;
  providerId: string;
}
```

read/write 経路：
- `xl/persons/personN.xml` を `wb.threadedAuthors` (= Person[]) に
- `xl/threadedCommentsN.xml` を sheet 単位で読み込み

### 1.4 編集 API

```ts
export function addComment(ws: Worksheet, ref: string, author: string, text: string | RichText): void;
export function getComments(ws: Worksheet, ref: string): LegacyComment[];
export function removeComment(ws: Worksheet, ref: string): void;

export function addThreadedComment(ws: Worksheet, ref: string, person: Person, text: string): ThreadedComment;
export function replyToThreadedComment(ws: Worksheet, parent: ThreadedComment, person: Person, text: string): ThreadedComment;
```

### 1.5 受け入れ条件

- [ ] openpyxl `comments/tests/data/` の round-trip
- [ ] threaded comment + legacy comment が両方ある xlsx の round-trip
- [ ] 新規追加した legacy comment が VML drawing と整合している（位置 / サイズ）

---

## 2. ハイパーリンク

### 2.1 型

```ts
export interface Hyperlink {
  ref: string;               // "A1" or "A1:B2"
  /** 外部 URL の場合 */
  target?: string;
  /** 内部 (workbook 内) jump 先 */
  location?: string;
  display?: string;
  tooltip?: string;
  /** rels で割り当てられる id（外部 URL のみ） */
  rId?: string;
}
```

外部 URL は `worksheet.xml.rels` に登録される。内部リンクは `<hyperlink location="Sheet2!A1"/>`。

### 2.2 編集 API

```ts
export function setHyperlink(ws: Worksheet, ref: string, opts: { target?: string; location?: string; display?: string; tooltip?: string }): void;
export function clearHyperlink(ws: Worksheet, ref: string): void;
```

cell value とハイパーリンクは独立して管理する。Worksheet 上の `hyperlinks` 配列に追加し、書き出し時に `<hyperlinks>` 要素として出力。

### 2.3 受け入れ条件

- [ ] 外部 URL / 内部 location 両方の round-trip
- [ ] 200 個のハイパーリンクで rels.xml が正しく生成される
- [ ] tooltip / display の round-trip

---

## 3. テーブル（Excel Table）

### 3.1 型

```ts
// src/worksheet/table.ts
export interface TableDefinition {
  id: number;                  // tableN
  displayName: string;
  name?: string;
  ref: string;                 // "A1:E10"
  totalsRowShown?: boolean;
  totalsRowCount?: number;
  headerRowCount?: number;
  styleInfo?: TableStyleInfo;
  columns: TableColumn[];
  autoFilter?: AutoFilter;
  sortState?: SortState;
}

export interface TableColumn {
  id: number;
  name: string;
  totalsRowFunction?: 'sum' | 'min' | 'max' | 'count' | 'average' | 'stdDev' | 'var' | 'custom';
  totalsRowFormula?: string;
  totalsRowLabel?: string;
  calculatedColumnFormula?: string;
}

export interface TableStyleInfo {
  name?: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
}
```

### 3.2 read/write

- `xl/tables/tableN.xml` 自体は schema 経由
- 各 worksheet の rels から `tablePart` を解決
- `<tableParts>` を sheet XML 末尾に出力

### 3.3 編集 API

```ts
export function addTable(ws: Worksheet, opts: Pick<TableDefinition, 'displayName' | 'ref' | 'styleInfo'> & { columns?: string[] }): TableDefinition;
export function removeTable(ws: Worksheet, displayName: string): void;
export function getTable(ws: Worksheet, displayName: string): TableDefinition | undefined;
```

### 3.4 受け入れ条件

- [ ] 1 〜 5 個のテーブルがある xlsx の round-trip
- [ ] テーブル列名と displayName の正規化が openpyxl 等価
- [ ] テーブルスタイル参照（builtin / custom）の round-trip

---

## 4. AutoFilter / SortState

### 4.1 型

```ts
export interface AutoFilter {
  ref: string;
  filterColumns: FilterColumn[];
  sortState?: SortState;
}

export type FilterColumn =
  | { kind: 'filters'; colId: number; values: Array<string | number>; blank?: boolean }
  | { kind: 'customFilters'; colId: number; and?: boolean; filters: CustomFilter[] }
  | { kind: 'top10'; colId: number; top: boolean; percent: boolean; val: number }
  | { kind: 'dynamicFilter'; colId: number; type: string; val?: number }
  | { kind: 'colorFilter'; colId: number; dxfId: number; cellColor: boolean }
  | { kind: 'iconFilter'; colId: number; iconSet: string; iconId: number };

export interface CustomFilter { operator: 'equal' | 'lessThan' | 'lessThanOrEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'notEqual'; val: string | number; }

export interface SortState {
  ref: string;
  caseSensitive?: boolean;
  columnSort?: boolean;
  sortConditions: SortCondition[];
}
export interface SortCondition {
  ref: string;
  descending?: boolean;
  sortBy?: 'value' | 'cellColor' | 'fontColor' | 'icon';
  customList?: string;
  dxfId?: number;
  iconSet?: string;
  iconId?: number;
}
```

### 4.2 受け入れ条件

- [ ] openpyxl `worksheet/tests/test_filters.py` 系のフィクスチャの round-trip

---

## 5. データ検証

### 5.1 型

```ts
export interface DataValidation {
  type: 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';
  operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual';
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  errorTitle?: string;
  error?: string;
  errorStyle?: 'stop' | 'warning' | 'information';
  promptTitle?: string;
  prompt?: string;
  showDropDown?: boolean;
  sqref: MultiCellRange;        // 適用範囲
}
```

### 5.2 編集 API

```ts
export function addDataValidation(ws: Worksheet, dv: DataValidation): void;
```

### 5.3 受け入れ条件

- [ ] list 型・date 型・custom 型の round-trip
- [ ] 大量レンジを表す compact sqref（`A1 A2 A3 …` の正規化）の生成

---

## 6. 条件付き書式

### 6.1 型

```ts
// src/formatting/conditional.ts
export interface ConditionalFormatting {
  sqref: MultiCellRange;
  rules: Rule[];
  pivot?: boolean;
}

export type Rule =
  | { kind: 'expression'; priority: number; stopIfTrue?: boolean; dxfId: number; formula: string[] }
  | { kind: 'cellIs'; priority: number; stopIfTrue?: boolean; dxfId: number; operator: CellIsOp; formula: string[] }
  | { kind: 'colorScale'; priority: number; cfvo: Cfvo[]; colors: Color[] }
  | { kind: 'dataBar'; priority: number; cfvo: Cfvo[]; color: Color; minLength?: number; maxLength?: number; showValue?: boolean }
  | { kind: 'iconSet'; priority: number; iconSet: IconSetType; cfvo: Cfvo[]; reverse?: boolean; showValue?: boolean }
  | { kind: 'top10'; priority: number; dxfId: number; rank: number; bottom?: boolean; percent?: boolean }
  | { kind: 'aboveAverage'; priority: number; dxfId: number; aboveAverage?: boolean; equalAverage?: boolean; stdDev?: number }
  | { kind: 'uniqueValues' | 'duplicateValues'; priority: number; dxfId: number }
  | { kind: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'; priority: number; dxfId: number; text: string; operator: 'containsText' | 'notContains' | 'beginsWith' | 'endsWith' }
  | { kind: 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors'; priority: number; dxfId: number }
  | { kind: 'timePeriod'; priority: number; dxfId: number; timePeriod: 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisMonth' | 'lastMonth' | 'nextMonth' | 'thisWeek' | 'lastWeek' | 'nextWeek' };

export type CellIsOp = 'lessThan' | 'lessThanOrEqual' | 'equal' | 'notEqual' | 'greaterThanOrEqual' | 'greaterThan' | 'between' | 'notBetween';

export interface Cfvo {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula';
  val?: string;
  gte?: boolean;
}

export type IconSetType = '3Arrows' | '3ArrowsGray' | '3Flags' | /* … 全 19 種を network */;
```

### 6.2 編集 API

```ts
export function addConditionalFormatting(ws: Worksheet, sqref: string, rule: Rule): void;
export function getConditionalFormatting(ws: Worksheet): ConditionalFormatting[];
```

### 6.3 受け入れ条件

- [ ] 全 rule kind の round-trip
- [ ] dxfId が Stylesheet.dxfs と整合（dedup 後の index）
- [ ] pivot 用のフォーマットも保持（pivot 内部までは追わない）

---

## 7. 定義名 / 印刷タイトル / 印刷範囲

`src/workbook/defined-name.ts`:

```ts
export interface DefinedName {
  name: string;
  comment?: string;
  customMenu?: string;
  description?: string;
  help?: string;
  statusBar?: string;
  /** scope: 0 = global, 1+ = sheet index */
  localSheetId?: number;
  hidden?: boolean;
  function?: boolean;
  vbProcedure?: boolean;
  xlm?: boolean;
  functionGroupId?: number;
  shortcutKey?: string;
  publishToServer?: boolean;
  workbookParameter?: boolean;
  /** 値 */
  value: string;            // formula / range / static value
}
```

特殊な事前定義名：
- `_xlnm.Print_Area` → `worksheet.printArea`
- `_xlnm.Print_Titles` → `worksheet.printTitleRows`/`printTitleCols`
- `_xlnm._FilterDatabase` → AutoFilter

これらは parser 側で `Worksheet` フィールドへ振り分け、書き戻し時に `_xlnm.*` として再生する。

### 7.1 編集 API

```ts
export function defineName(wb: Workbook, name: string, value: string, opts?: { localSheetId?: number; hidden?: boolean; comment?: string }): void;
export function removeName(wb: Workbook, name: string): void;
export function listNames(wb: Workbook, opts?: { scope?: 'global' | number }): DefinedName[];
```

### 7.2 受け入れ条件

- [ ] global / sheet-scoped の round-trip
- [ ] `Print_Area`, `Print_Titles` が `worksheet.printArea`/`printTitleRows` に分離されて再生も同じ XML
- [ ] 名前の検証（Excel の予約語、空白、最大長 255）が `defineName` 時にチェックされる

---

## 8. ヘッダ・フッタ / 印刷設定 / ページ区切り

フェーズ3 で passthrough にしていたものを完全構造化：

```ts
export interface PageSetup {
  orientation?: 'portrait' | 'landscape';
  paperSize?: number;     // 1 = letter, 9 = A4, …
  scale?: number;         // %
  fitToHeight?: number;
  fitToWidth?: number;
  pageOrder?: 'downThenOver' | 'overThenDown';
  blackAndWhite?: boolean;
  draft?: boolean;
  cellComments?: 'none' | 'asDisplayed' | 'atEnd';
  errors?: 'displayed' | 'blank' | 'dash' | 'NA';
  horizontalDpi?: number;
  verticalDpi?: number;
  copies?: number;
  firstPageNumber?: number;
  useFirstPageNumber?: boolean;
}

export interface HeaderFooter {
  oddHeader?: HeaderFooterParts;
  oddFooter?: HeaderFooterParts;
  evenHeader?: HeaderFooterParts;
  evenFooter?: HeaderFooterParts;
  firstHeader?: HeaderFooterParts;
  firstFooter?: HeaderFooterParts;
  differentOddEven?: boolean;
  differentFirst?: boolean;
  alignWithMargins?: boolean;
  scaleWithDoc?: boolean;
}

export interface HeaderFooterParts {
  left?: string;
  center?: string;
  right?: string;
}
```

`HeaderFooterParts` の文字列は Excel の `&P`, `&D`, `&"Arial,Bold"` 等のエスケープを保持。helper でパース／生成する：

```ts
export function parseHeaderFooterText(s: string): HeaderFooterToken[];
export function buildHeaderFooterText(tokens: HeaderFooterToken[]): string;
```

### 8.1 受け入れ条件

- [ ] 横位置・縦位置・スケール・印刷タイトルの round-trip
- [ ] header / footer のフォント指定（`&"Calibri,Bold"`）の解析と再生

---

## 9. 外部リンク

`src/workbook/external-link/`:

read 時にバイナリ + 構造を保存し、書き戻すだけ（編集 API は提供しない）。

```ts
export interface ExternalLink {
  rId: string;
  target: string;
  bookXml: Uint8Array;   // 元 XML をそのまま保持
  bookRels: Uint8Array;
}
```

### 9.1 受け入れ条件

- [ ] 外部リンク付き xlsx の round-trip で `xl/externalLinks/*` が消えない
- [ ] 編集 API は意図的に未提供（ドキュメントで明示）

---

## 10. 完了条件（フェーズ5 全体）

- [ ] §1〜§9 の各受け入れ条件
- [ ] フェーズ3 の passthrough 経路が正しく構造化に置き換わっている
- [ ] バンドルサイズ予算違反なし
- [ ] ロード／セーブのベンチマーク regression なし
