# 06. フェーズ4: ストリーミング（read-only / write-only）

**目的**: 数百万セル規模の xlsx を OOM なしで処理する。
**期間目安**: 3週間
**前提**: フェーズ3 完了
**完了条件**: 100 万行 × 30 列の xlsx を 1GB 以下のヒープで read / write できる。

## 1. 全体方針

openpyxl は `read_only=True` / `write_only=True` の二系統を持つ：
- `read_only`: archive を保持し、`worksheet.iter_rows()` 呼び出し時に SAX で sheet を再パース。
- `write_only`: temp file に worksheet XML を書き、最後に zip に貼る。

TS では **Web Streams API** が両環境で使えることを利用し、temp file は使わない。worksheet XML を ZIP の deflate stream に **直接** 流す。

サブパスで分離（[01-architecture.md](./01-architecture.md) §6）：
```
import { loadWorkbookStream } from 'ooxml-js/streaming';
import { createWriteOnlyWorkbook } from 'ooxml-js/streaming';
```

## 2. read-only

### 2.1 API

```ts
// src/streaming/read-only-workbook.ts
export async function loadWorkbookStream(source: XlsxSource, opts?: LoadOptions): Promise<ReadOnlyWorkbook>;

export interface ReadOnlyWorkbook {
  sheetNames: string[];
  properties: CoreProperties;
  /** worksheet を lazy にオープン */
  openWorksheet(name: string): ReadOnlyWorksheet;
  close(): Promise<void>;
}

export interface ReadOnlyWorksheet {
  title: string;
  /** dimension が信用できる場合のみ */
  maxRow?: number;
  maxCol?: number;
  iterRows(opts?: IterRowsOptions): AsyncIterableIterator<ReadOnlyCell[]>;
  /** 値だけ欲しい時の高速ルート */
  iterValues(opts?: IterRowsOptions): AsyncIterableIterator<CellValue[]>;
}

export interface IterRowsOptions {
  minRow?: number;
  maxRow?: number;
  minCol?: number;
  maxCol?: number;
}
```

### 2.2 実装ポイント

- `loadWorkbookStream` は workbook.xml / styles.xml / sharedStrings.xml の **メタだけ** 読み出して保持する（数百万 cell 環境でもメタは小さい）。
- **sharedStrings は `Uint32Array` index + 1 本の連結文字列** で持つ（メモリ効率）。
- `openWorksheet(name)` は `xl/worksheets/sheetN.xml` の bytes をオンデマンドで取得し、SAX iterator を返す。
- 並行して複数 sheet を iter する場合、ZIP の同一エントリを重ねて読まないよう entry buffer を再利用する。
- `iterRows` は **生成器** とし、各 row を `ReadOnlyCell[]` で返す。styleId は持たせるが、値型の解釈はオプション（`opts.parseDates: false` で serial のまま）。

### 2.3 ReadOnlyCell

```ts
export interface ReadOnlyCell {
  readonly row: number;
  readonly col: number;
  readonly value: CellValue;
  readonly styleId: number;
  /** lazy: getter で formatCode を引く */
  readonly numberFormat: string;
}
```

`numberFormat` は getter にし、必要時のみ Stylesheet を引く。Object.defineProperty を **moduleレベルで 1 回だけ定義** したテンプレートを `Object.create` で量産する（V8 の hidden class 共有）。

### 2.4 受け入れ条件

- [ ] 100 万行 × 30 列の xlsx を 1GB ヒープ以下で全行 iter pass
- [ ] iterRows の throughput が > 500k cells/s（M1）
- [ ] 並行して 2 sheets 同時 iter しても結果が混線しない
- [ ] `close()` で archive のハンドルが解放される（Node の fs handle / ブラウザの Blob ストリーム）

## 3. write-only

### 3.1 API

```ts
// src/streaming/write-only-workbook.ts
export async function createWriteOnlyWorkbook(sink: XlsxSink, opts?: WriteOnlyOptions): Promise<WriteOnlyWorkbook>;

export interface WriteOnlyOptions {
  /** 既存の値オブジェクトをそのまま渡せる */
  properties?: Partial<CoreProperties>;
  /** 書き出し中の最大 row 数（dimension タグに使う） */
  estimatedMaxRow?: number;
}

export interface WriteOnlyWorkbook {
  /** Worksheet を 1 個ずつ追加。前の ws が finalize されてから次が開ける（ZIP は順次しか書けないため） */
  addWorksheet(title: string): Promise<WriteOnlyWorksheet>;
  setProperty<K extends keyof CoreProperties>(k: K, v: CoreProperties[K]): void;
  defineNamedStyle(name: string, font?: Font, /* … */): void;
  /** 全ての ws.close() を呼んだ後、最後にこれを呼ぶ */
  finalize(): Promise<void>;
}

export interface WriteOnlyWorksheet {
  title: string;
  /** 行を 1 個追加 */
  appendRow(row: WriteOnlyRowItem[]): Promise<void>;
  /** 列幅・行高など事前設定 */
  setColumnWidth(col: number, width: number): void;
  /** 完了 */
  close(): Promise<void>;
}

export type WriteOnlyRowItem =
  | CellValue
  | { value: CellValue; style?: WriteOnlyStyle };

export interface WriteOnlyStyle {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  numberFormat?: string;
  protection?: Protection;
}
```

### 3.2 実装ポイント

- `createWriteOnlyWorkbook` は ZIP writer（streaming）、Stylesheet（プール）、SharedStringAccumulator を初期化する。
- `addWorksheet` は **直前の ws が close されているか** をチェックする（ZIP は同時書き込み不可）。
  - 並列に append したい場合は **temp Buffer 経由**を使うオプションを別途用意（`addWorksheet({ buffered: true })`）。temp Buffer は WeakRef で解放可能。
- `appendRow` は worksheet の `<sheetData>` 要素に対して row 1 個分の XML を直接 stream に書く。
- `appendRow` 経由の値はその場で Stylesheet にスタイル登録。dedup は in-memory プールで O(1)。
- styles.xml / sharedStrings.xml は **最後に書き出す** ため、worksheet のあとに ZIP entry を追加する。openpyxl は最後にやる。
- finalize で：
  1. すべての worksheet の close を確認
  2. Stylesheet を `xl/styles.xml` にシリアライズ
  3. SharedStrings を `xl/sharedStrings.xml` にシリアライズ
  4. workbook.xml を出力
  5. Manifest / rels を出力
  6. ZIP finalize

### 3.3 sheetData 出力ホットパス

```ts
// src/streaming/sheet-data-writer.ts
export interface SheetDataWriter {
  startRow(rowIdx: number, height?: number, customHeight?: boolean): void;
  cell(coord: string, value: CellValue, styleId: number): void;
  endRow(): void;
  finish(): void;
}
```

内部実装：
- `Uint8Array` チャンクのリングバッファ（既定 64KB）に書き出す。
- バッファ満杯になったら ZIP writer に flush（backpressure を尊重）。
- 文字列セルは shared strings accumulator に dedup add → index を `<v>` に書く。
- 数値セルは `Number.toString()` を直接書く（`toFixed` などはユーザ責任）。
- dataType 判定の if 連鎖を最小化（`typeof v` と `Number.isFinite` だけで分岐）。

### 3.4 受け入れ条件

- [ ] 100 万行 × 30 列を < 10s / 1GB ヒープ未満で write
- [ ] write-only モードで Stylesheet の dedup が壊れない
- [ ] `appendRow` の入力に `{ value, style }` を混ぜると正しい xfId に解決される
- [ ] 出力 xlsx を Excel 365 で開いて壊れていない（自動 QA: headless で xlsx-validator にかける）

## 4. 共通の品質ゲート

- [ ] 通常モードと streaming モードで read 結果が論理的に同一（cell 値・styleId）
- [ ] `vitest bench` で性能リグレッション 25% 以内
- [ ] バンドルサイズ予算: `ooxml-js/streaming` ≤ 80KB min+gz

## 5. 実装上の注意

- `WritableStream` / `ReadableStream` の cancel/abort 時にも temp 状態が漏れないこと（`AbortSignal` で finalize スキップ + cleanup）。
- ブラウザの File API は `FileSystemWritableFileStream` をネイティブで提供する環境がある（Origin Private File System）。`XlsxSink` の helper を後で拡張可能にしておく。
- Node の `node:fs/promises` の `FileHandle.writableWebStream()` は実験的。フォールバックとして Node の `Writable` も sink helper に残す。
