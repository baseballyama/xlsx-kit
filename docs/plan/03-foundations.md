# 03. フェーズ1: 基盤層

**目的**: 全フェーズが乗る土台を作る。ZIP・XML・I/O 抽象・Schema・パッケージング層を整備する。
**期間目安**: 4〜5週間
**前提**: なし
**完了条件**: 後続フェーズが着手できる状態。本フェーズ単独でラウンドトリップは未達でよい。

## 1. I/O 抽象（`src/io/`）

### 1.1 Source / Sink

```ts
// src/io/source.ts
export interface XlsxSource {
  /** 全バイトをメモリに読み出す */
  toBytes(): Promise<Uint8Array>;
  /** 順次読み出すストリーム（fallback として toBytes 可） */
  toStream?(): ReadableStream<Uint8Array>;
}

// src/io/sink.ts
export interface XlsxSink {
  /** メモリにバッファして最後に flush する */
  toBytes?(): { write(chunk: Uint8Array): void; finish(): Promise<Uint8Array> };
  /** ストリーム書き出し */
  toStream?(): WritableStream<Uint8Array>;
}
```

### 1.2 環境別の helper

`src/io/node.ts`:

```ts
export function fromFile(path: string): XlsxSource;
export function fromBuffer(buf: Buffer | Uint8Array): XlsxSource;
export function fromReadable(stream: import('node:stream').Readable): XlsxSource;

export function toFile(path: string): XlsxSink;
export function toBuffer(): XlsxSink & { result(): Buffer };
export function toWritable(stream: import('node:stream').Writable): XlsxSink;
```

`src/io/browser.ts`:

```ts
export function fromBlob(blob: Blob): XlsxSource;
export function fromFile(file: File): XlsxSource;        // 同名で OK
export function fromArrayBuffer(buf: ArrayBuffer | Uint8Array): XlsxSource;
export function fromResponse(res: Response): XlsxSource; // fetch 直結

export function toBlob(mime?: string): XlsxSink & { result(): Blob };
export function toArrayBuffer(): XlsxSink & { result(): ArrayBuffer };
```

`package.json` の `exports` 条件でこれらを `openxml-js/io` のサブパスから両環境が同名 import できるようにする。Node 専用の `fromFile(path)` とブラウザ専用の `fromFile(File)` は **シグネチャ重複でも実体は別ファイル**で ok。

### 1.3 受け入れ条件

- [ ] Node の Buffer / Readable / file path でそれぞれ source/sink が成り立つ
- [ ] ブラウザの File / Blob / ArrayBuffer / Response でそれぞれ source/sink が成り立つ
- [ ] 0 byte / 8GB のダミーストリームでも throw せず、エラーが起きた場合は `OpenXmlIoError` で投げる
- [ ] `Source` / `Sink` 型が同一インターフェースを共有し、上位レイヤから環境差を意識しない

## 2. ZIP 層（`src/zip/`）

### 2.1 Reader

fflate の Unzip を使い、**ファイル全体を一気に展開しない**。エントリ単位で取り出せるようにする。

```ts
// src/zip/reader.ts
export interface ZipEntry {
  path: string;
  bytes: Uint8Array;        // 解凍済み
}

export async function openZip(source: XlsxSource): Promise<{
  list(): string[];
  read(path: string): Uint8Array;          // 同期（メモリ展開済の場合）
  readAsync(path: string): Promise<Uint8Array>; // 大規模時の lazy
  close(): void;
}>;
```

実装方針：
- 小〜中規模（< 50 MB を目安）はメモリ展開（`fflate.unzipSync` / `fflate.unzip`）。
- それ以上は **stream-based unzip**（`fflate.Unzip` クラス）でエントリのオフセットを覚え、必要時にだけ解凍する LRU キャッシュ。
- `XlsxSource.toBytes()` 経由で得たバッファは再利用する。

### 2.2 Writer

fflate の `Zip` （streaming）を使う。エントリごとに `ZipPassThrough`（無圧縮）または `ZipDeflate`（圧縮）を選び、ストリーム的に sink に流す。

```ts
// src/zip/writer.ts
export interface ZipWriter {
  addEntry(path: string, bytes: Uint8Array | ReadableStream<Uint8Array>, opts?: { compress?: boolean }): Promise<void>;
  finalize(): Promise<void>;
}

export function createZipWriter(sink: XlsxSink): ZipWriter;
```

OOXML エントリは原則 deflate 圧縮。ただし以下は **無圧縮**：
- `xl/media/*`（既に PNG/JPEG 等で圧縮済み）
- `xl/vbaProject.bin`（既に圧縮済み）

### 2.3 受け入れ条件

- [ ] openpyxl が出力した xlsx を読み込み、すべてのエントリ path/size がオリジナルと一致
- [ ] 自分が出力した xlsx を再度読めて、エントリ集合が同型
- [ ] 100MB 級のフィクスチャで OOM しない（streaming 経路）
- [ ] ZIP64 を読める（テスト: openpyxl の bigfoot.xlsx を読める）
- [ ] ZIP64 を必要に応じて書ける（4GB 越えの巨大単一エントリは fail で良い、エラーを明示）

## 3. XML 層（`src/xml/`）

### 3.1 二系統の API

| 用途 | 実装 | 採用 lib |
|------|------|---------|
| 小さい XML（styles, workbook, manifest 等）の DOM 解析 | `parseXml(bytes): XmlNode` | fast-xml-parser |
| シリアライズ全般（write） | `serializeXml(node): Uint8Array` | 自作（テンプレート文字列ベース） |
| 大きい XML（worksheet sheetData, sharedStrings）の SAX 解析 | `iterParse(stream): AsyncIterableIterator<SaxEvent>` | saxes |
| ストリーム書き込み | `XmlStreamWriter` | 自作（後述 §5） |

### 3.2 内部 XML ノード表現

軽量 plain object。DOM ライブラリは使わない（速度・size のため）。

```ts
// src/xml/tree.ts
export interface XmlNode {
  /** 名前空間つき QName: '{http://schemas.../}tagname' or 'tagname' */
  name: string;
  attrs: Record<string, string | undefined>;
  /** テキストノード（複数 child の場合は children に Text を入れる） */
  text?: string;
  children: XmlNode[];
}
export function el(name: string, attrs?: Record<string, string>, children?: XmlNode[]): XmlNode;
```

`name` は `'{ns}local'` 形式で保持する。openpyxl も同じ。書き出し時に prefix を付け直す。

### 3.3 namespace 定数

`src/xml/namespaces.ts` に OOXML の全 namespace を集約する。openpyxl の `xml/constants.py` を完全に網羅する：

```ts
export const SHEET_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const REL_NS        = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const PKG_REL_NS    = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const CONTYPES_NS   = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const DRAWING_NS    = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const CHART_NS      = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
export const SS_DRAW_NS    = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
export const PIC_NS        = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
export const DC_NS         = 'http://purl.org/dc/elements/1.1/';
export const DCTERMS_NS    = 'http://purl.org/dc/terms/';
export const DCMITYPE_NS   = 'http://purl.org/dc/dcmitype/';
export const XSI_NS        = 'http://www.w3.org/2001/XMLSchema-instance';
export const XML_NS        = 'http://www.w3.org/XML/1998/namespace';
export const COMMENTS_NS   = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
// ほか mc/x14/x15/x16 の MS 拡張も列挙
```

prefix マッピングは write 時に **固定 prefix** を採用（openpyxl 互換）。例: `xmlns:r="…relationships"`, `xmlns:xdr="…spreadsheetDrawing"`, `xmlns:a="…drawingml/main"`, `xmlns:c="…drawingml/chart"`。

### 3.4 セキュリティ

- `parseXml` は **DTD 参照、外部実体（XXE）、外部 DOCTYPE をすべて拒否** する設定で `fast-xml-parser` を呼ぶ。
- `iterParse` (saxes) も同様に external entity 解決を無効化。
- 読み込んだ XML に `<!DOCTYPE` が含まれていたら **明示的に reject**（`OpenXmlSchemaError('DTD declarations are not permitted in OOXML payloads')`）。

### 3.5 受け入れ条件

- [ ] openpyxl テストフィクスチャの XML 群すべてを `parseXml` → `serializeXml` で round-trip し、canonical-XML 比較で一致
- [ ] `iterParse` で 100 万行の sheetData を OOM なし、< 10s で消費（M1 等の通常開発機）
- [ ] XXE / DTD 注入のテストケースが reject される

## 4. Schema 層（`src/schema/`）

openpyxl の `descriptors/` を **クラスではなく純データの Schema + 純粋関数** に置き換える。

### 4.1 Schema 型

```ts
// src/schema/core.ts
export type Primitive = 'string' | 'int' | 'float' | 'bool' | 'datetime';

export interface AttrDef {
  kind: Primitive | 'enum';
  values?: readonly string[];     // enum の選択肢
  optional?: boolean;
  default?: unknown;
  min?: number; max?: number;     // MinMax
  pattern?: RegExp;
  /** key 名と異なる XML attribute 名を使う場合 */
  xmlName?: string;
  /** 属性の名前空間（{ns}attr 形式で保存される） */
  xmlNs?: string;
}

export type ElementDef =
  | { kind: 'text'; name: string; xmlNs?: string; primitive: Primitive; optional?: boolean }
  | { kind: 'object'; name: string; xmlNs?: string; schema: () => Schema<unknown>; optional?: boolean }
  | { kind: 'sequence'; name: string; xmlNs?: string; itemSchema: () => Schema<unknown>; container?: { tag: string; count?: boolean } }
  | { kind: 'union'; name: string; variants: Array<{ tag: string; schema: () => Schema<unknown> }> }
  | { kind: 'empty'; name: string; xmlNs?: string; /** EmptyTag 相当: 存在で true、不在で false */ };

export interface Schema<T> {
  tagname: string;
  xmlNs?: string;
  attrs: Record<keyof T & string, AttrDef> | {};
  elements: ElementDef[];
  /** XML 上の出力順（指定がなければ elements の順） */
  order?: string[];
  /** カスタム前後処理（読み込み後の正規化や互換ハック用） */
  postParse?: (value: T, node: XmlNode) => T;
  preSerialize?: (value: T) => T;
}

export function defineSchema<T>(s: Schema<T>): Schema<T> { return s; }
```

### 4.2 シリアライザ / デシリアライザ

```ts
// src/schema/serialize.ts
export function toTree<T>(value: T, schema: Schema<T>): XmlNode;
export function fromTree<T>(node: XmlNode, schema: Schema<T>): T;
```

実装方針：
- すべて純粋関数。class メソッドではない。
- 内部は **switch on `kind`**。再帰呼び出しで子 schema を解決。
- 値は plain object。`new` しない。
- 必須 attr が欠落していたら `OpenXmlSchemaError` を投げる（読み込み時）。
- 未対応の attr/element は警告ログ + `passthroughExt` フィールド（JSON 化された XML フラグメント）に保存して書き戻す（`extLst` 系の前方互換用途）。

### 4.3 Schema レジストリと cycle

`Border` の Side や Stylesheet の循環参照は **lazy schema getter** (`schema: () => SideSchema`) で解決する。

### 4.4 Validation 戦略

- **読み込み時**: schema で型/enum/MinMax を検証。失敗時は `OpenXmlSchemaError`。
- **書き出し時**: 検証は最小限。型シグネチャに従っていることを TS 型システムで担保する前提。
- **runtime validation の重さがホットパスを汚さない**ように、worksheet `cell` の値は schema を経由しない（直接 `<c>` を文字列テンプレートで吐く）。スタイル系の小さな構造のみ schema 経由。

### 4.5 受け入れ条件

- [ ] `Font`, `Border`, `Side`, `Alignment`, `Protection`, `Color`, `NumberFormat` を schema 化して round-trip テスト pass
- [ ] schema 1 件の最小 module サイズ（min+gz）を計測し、≤ 1.5KB / schema を維持
- [ ] `defineSchema` が tree-shake で不要な schema を除去できる（テスト: ダミー entry で `import { foo } from 'openxml-js/styles'` 後の bundle に未使用 schema が含まれないこと）

## 5. XmlStreamWriter（`src/xml/stream-writer.ts`）

openpyxl の `et_xmlfile.xmlfile` は generator + context manager の二段構えだが、TS では state machine + chunk emit の単純なクラス**ではない**関数群で実装する。

```ts
export interface XmlStreamWriter {
  /** 開始要素 */
  start(name: string, attrs?: Record<string, string>): void;
  /** テキストノード */
  text(s: string): void;
  /** 子要素全体（DOM ノード） */
  writeNode(n: XmlNode): void;
  /** 終了要素 */
  end(): void;
  /** finalize して全バイトを出力 */
  finalize(): Uint8Array | Promise<void>;
}

export function createXmlStreamWriter(target: WritableStream<Uint8Array> | { append(chunk: Uint8Array): void }): XmlStreamWriter;
```

実装：
- 内部は `Uint8Array` のチャンク配列を溜める。
- 必要に応じて `WritableStream` に backpressure 対応で書き出す。
- 属性のエスケープ・テキストのエスケープは内蔵。`&`, `<`, `>`, `"`, `'` を順に置換。
- 名前空間は writer 生成時に prefix map を渡す（`{ '': SHEET_MAIN_NS, r: REL_NS }`）。

ホットパス（`<row>`/`<c>` の量産）は `writeNode` ではなく直接の文字列テンプレートで書く。`XmlStreamWriter` は構造化が必要な箇所に限定する。

### 5.1 受け入れ条件

- [ ] 100 万 cell の sheet を < 5s で出力（MacBook M1 想定）
- [ ] 100 万 cell の sheet を出力中、ヒープ使用が 200MB を超えない
- [ ] ストリーム書き出し中に `await` を挟んでも順序が崩れない

## 6. パッケージング層（`src/packaging/`）

### 6.1 Manifest（[Content_Types].xml）

```ts
// src/packaging/manifest.ts
export interface Manifest {
  defaults: Array<{ ext: string; contentType: string }>;
  overrides: Array<{ partName: string; contentType: string }>;
}

export function makeManifest(): Manifest;
export function addDefault(m: Manifest, ext: string, contentType: string): void;
export function addOverride(m: Manifest, partName: string, contentType: string): void;
export function manifestToBytes(m: Manifest): Uint8Array;
export function manifestFromBytes(b: Uint8Array): Manifest;
```

dedup ロジック付き。`addDefault` / `addOverride` は同じエントリの追加を no-op。

### 6.2 Relationships（*.rels）

```ts
export interface Relationship {
  id: string;       // rId#
  type: string;     // 完全 URI
  target: string;   // 相対パス or URL
  targetMode?: 'External' | 'Internal';
}
export interface Relationships { rels: Relationship[]; }

export function makeRelationships(): Relationships;
export function appendRel(rels: Relationships, type: string, target: string, targetMode?: 'External'): Relationship;
export function findByType(rels: Relationships, type: string): Relationship | undefined;
export function relsToBytes(rels: Relationships): Uint8Array;
export function relsFromBytes(b: Uint8Array): Relationships;
```

`appendRel` は最小未使用 rId を割り当てる（openpyxl `relationship.py:59-62`）。

### 6.3 Document properties

`docProps/core.xml`, `docProps/app.xml`, `docProps/custom.xml` をそれぞれ schema 化：

- `CoreProperties`: title, subject, creator, keywords, description, lastModifiedBy, modified, created, category, contentStatus, …
- `ExtendedProperties`: Application, AppVersion, Company, …
- `CustomProperties`: 複数 `CustomDocumentProperty`（Name, fmtid, pid, value）

これらは [02-mapping.md](./02-mapping.md) の対応表通り。

### 6.4 受け入れ条件

- [ ] openpyxl が出力した [Content_Types].xml と等価な manifest を生成できる
- [ ] rId 衝突の起こりえる順序で append しても正しく一意になる
- [ ] custom doc props の round-trip が openpyxl の `tests/test_read_write_custom_doc_props.py` 相当のフィクスチャで一致する

## 7. utils（`src/utils/`）

| ファイル | 内容 |
|---------|------|
| `coordinate.ts` | `columnLetterFromIndex(n)`, `columnIndexFromLetter(s)`, `coordinateFromString(s) → [col, row]`, `rangeBoundaries(s)`, `rangeToTuple(s)` |
| `datetime.ts` | `excelToDate(serial, epoch?)`, `dateToExcel(date, epoch?)`, `excelToDuration(serial)`, `durationToExcel(ms)`, ISO 8601 helper |
| `units.ts` | `EMU = 9525` 等の定数、`emuFromPixel(px)`, `pxFromEmu(emu)` |
| `inference.ts` | `inferCellType(value)`: `'n' | 's' | 'b' | 'd' | 'f' | 'e'` |
| `escape.ts` | XML エスケープ（attr / text 別）、ascii control char の制御 |
| `exceptions.ts` | `OpenXmlError` 階層 |

### 7.1 coordinate

openpyxl の `utils/cell.py` のロジックを移植。`columnIndexFromLetter` は **lru_cache** 相当の `Map<string, number>` キャッシュをモジュール内で持つ（Worksheet 走査時に同じ列文字を何度も解析するため）。最大列 16384 (XFD) の制約を入れる。

### 7.2 datetime

- 1900 epoch（既定）/ 1904 epoch を引数で切り替え。
- 1900 leap-year bug（serial 60 = 1900-02-29 を Excel が扱う）を補正。
- `Date` を直接使うが、ホットパス（worksheet read で 100 万件） では `serial` のまま保持し、ユーザが要求時に変換する。

### 7.3 受け入れ条件

- [ ] coordinate 系は openpyxl の単体テスト全件を移植して pass
- [ ] datetime は 1900/1904 両方で round-trip 一致
- [ ] inference は openpyxl の `cell.py:_TYPES` と同一の判定
- [ ] エラー階層は呼び出し側で discriminated union として扱える

## 8. compat（`src/compat/`）

| ファイル | 内容 |
|---------|------|
| `numbers.ts` | `isFiniteNumber(x)`, `isInteger(x)`, optional decimal helper（dynamic import） |
| `singleton.ts` | （不要なら削除可。`Object.freeze` で代替） |

最小限。Python の dynamic typing 補助が中心なので TS では不要なものが多い。

## 9. テスト（`tests/phase-1/`）

- `tests/phase-1/zip.test.ts`: ZIP read/write の round-trip。
- `tests/phase-1/xml.test.ts`: parse/serialize round-trip。
- `tests/phase-1/iterparse.test.ts`: SAX 経由の cell 数カウントが正しい。
- `tests/phase-1/schema.test.ts`: 簡易な Schema（Border, Side）の round-trip。
- `tests/phase-1/packaging.test.ts`: Manifest + Relationships 生成の golden。
- `tests/phase-1/coordinate.test.ts`, `datetime.test.ts`, `inference.test.ts`.
- `tests/phase-1/io-node.test.ts` / `io-browser.test.ts`: 各 source/sink がインタフェース通り動く。

## 10. 完了条件（フェーズ1 全体）

すべて pass で次フェーズへ：

- [ ] §1〜§8 の各受け入れ条件
- [ ] CI の typecheck / lint / test / size budget が通る
- [ ] `src/index.ts` から低レベル API（`openZip`, `parseXml`, `toTree`, `fromTree`, `manifestToBytes` 等）が export されている
- [ ] サンプル: 「openpyxl が作った最小 xlsx を解凍 → manifest を読む → 元と等価な manifest を出力 → 再ジップ」が end-to-end で動く（**ただしまだ Workbook データ層は未実装**）
