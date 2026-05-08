# 05. フェーズ3: read / write 実装

**目的**: フェーズ2 で作ったコアモデルと、フェーズ1 で作った XML/ZIP/Schema 層を繋ぎ、xlsx ファイルの **読み書きラウンドトリップ** を成立させる。
**期間目安**: 5〜6週間
**前提**: フェーズ1, 2 が完了
**完了条件**: openpyxl が出力する基本的な xlsx を `loadWorkbook → 編集 → saveWorkbook` で正しく round-trip できる。

## 1. 全体フロー

### 1.1 loadWorkbook

参照: openpyxl `reader/excel.py:316-349`。

```ts
// src/public/load.ts
export async function loadWorkbook(source: XlsxSource, opts?: LoadOptions): Promise<Workbook>;

export interface LoadOptions {
  readOnly?: boolean;          // → ReadOnlyWorkbook を返す（フェーズ4）
  keepLinks?: boolean;         // 外部リンクを保持
  keepVba?: boolean;           // VBA バイナリを passthrough
  dataOnly?: boolean;          // 数式の cached value のみ採用、formula 文字列を捨てる
  richText?: boolean;          // inline string を rich text として読む
  logger?: Logger;
}
```

実装ステップ（順序固定）:
1. `openZip(source)` で archive を取得
2. `[Content_Types].xml` を解析 → Manifest
3. Manifest から workbook part path を解決（XLSX/XLSM/XLTX/XLTM の content type を見る）
4. `xl/_rels/workbook.xml.rels` を解析 → Workbook の rels マップ
5. `xl/sharedStrings.xml`（あれば）を SAX で stream parse → `wb.sharedStrings`
6. `xl/styles.xml` を解析 → `wb.styles`
7. `xl/theme/theme1.xml`（あれば）を **バイナリのまま** `wb.themeXml` に保存
8. `docProps/core.xml`, `docProps/app.xml`, `docProps/custom.xml` を解析
9. `xl/workbook.xml` を解析 → sheets リスト・defined names・external refs・bookViews・calcPr
10. 各 worksheet について：
    - `xl/worksheets/sheetN.xml.rels` を読む
    - 通常モードなら sheet XML を SAX iterparse で full load → `Worksheet`
    - read-only モードなら lazy 化（フェーズ4）
    - rels を解決して comments / drawings / tables / hyperlinks / pivot ref / vmlDrawing を読む
11. chartsheet があれば同様に読む
12. external links（読み込みのみ。編集 API は提供しない）
13. VBA project/binary を `wb.vbaProject` に passthrough
14. archive を close（または read-only 用に保持）

### 1.2 saveWorkbook

参照: openpyxl `writer/excel.py:53-94`。

```ts
// src/public/save.ts
export async function saveWorkbook(wb: Workbook, sink: XlsxSink, opts?: SaveOptions): Promise<void>;

export interface SaveOptions {
  /** ZIP のデフォルト圧縮レベル (0-9) */
  compressionLevel?: number;
}
```

実装ステップ（順序固定）：
1. `wb.properties.modified` を現在時刻に更新（オプションで disable）
2. ZIP writer 作成（streaming）
3. `docProps/app.xml` を schema 経由でシリアライズ
4. `docProps/core.xml`
5. `xl/theme/theme1.xml` を passthrough
6. `docProps/custom.xml`（あれば）
7. **Worksheets**:
   - 通常 mode: 各 worksheet を SAX-style でストリーム生成（`XmlStreamWriter`）
   - write-only mode: 既にユーザがストリームしている内容を flush（フェーズ4）
   - 各 worksheet の rels（drawing, comments, tables, hyperlinks）を出力
8. **Chartsheets**
9. **Drawings**: `xl/drawings/drawingN.xml` + rels（フェーズ6）
10. **Charts**: `xl/charts/chartN.xml`（フェーズ6）
11. **Images**: `xl/media/imageN.xxx`（フェーズ6）
12. **Comments**: `xl/commentsN.xml` + `xl/drawings/vmlDrawingN.vml`（フェーズ5）
13. **External links**（フェーズ5）
14. **Tables**（フェーズ5）
15. `xl/sharedStrings.xml`（書き込み中に蓄積されたものを最後に flush）
16. `xl/styles.xml`
17. `xl/workbook.xml` + `xl/_rels/workbook.xml.rels`
18. `_rels/.rels`
19. VBA passthrough（必要なら）
20. `[Content_Types].xml` を最後に出力（全 part が確定した後）
21. `finalize()`

> 注：sharedStrings は worksheet 書き出し中に **新規文字列が出現** するため、最後に出力する必要がある。openpyxl と同じ。

## 2. workbook.xml の read/write

### 2.1 関連 schema

`src/workbook/package.ts` で `WorkbookPackage` schema を定義する（openpyxl `packaging/workbook.py:91-186`）：

| 要素 | TS フィールド |
|------|-------------|
| `fileVersion` | `fileVersion?: FileVersion` |
| `workbookPr` | `workbookPr: WorkbookProperties` |
| `bookViews` | `bookViews: BookView[]` |
| `sheets` | `sheets: ChildSheet[]` |
| `definedNames` | `definedNames: DefinedName[]` |
| `calcPr` | `calcPr: CalcProperties` |
| `customWorkbookViews` | `customWorkbookViews: CustomWorkbookView[]` |
| `pivotCaches` | `pivotCaches: PivotCacheRef[]` |
| `externalReferences` | `externalReferences: ExternalReference[]` |
| `webPublishing` | `webPublishing?: WebPublishing` |
| `fileRecoveryPr` | `fileRecoveryPr?: FileRecoveryProperties[]` |
| `smartTagPr` | `smartTagPr?: SmartTagProperties` |
| `smartTagTypes` | `smartTagTypes?: SmartTagType[]` |
| `functionGroups` | `functionGroups?: FunctionGroups` |
| `extLst` | `extLst?: { passthroughXml: string }` |

### 2.2 sheets リストと rels の対応

`<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>` の `r:id` を `xl/_rels/workbook.xml.rels` から逆引きして、対応する `worksheets/sheetN.xml` を解決する。

write 時には逆向き：sheet ごとに rId を割り当て、target を `worksheets/sheetN.xml` 形式で出力。

### 2.3 受け入れ条件

- [ ] openpyxl の `tests/data/genuine/empty-with-styles.xlsx` を読んで再書き込み、canonical XML 等価
- [ ] `tests/data/genuine/sample.xlsx` で同上
- [ ] sheet を増やしたあとの save / load 回帰がない
- [ ] defined name の round-trip（global / sheet-scoped, print_area, print_titles）

## 3. styles.xml の read/write

参照: openpyxl `styles/stylesheet.py`。

write は **`Stylesheet` を schema → XML** で吐く。フィールドは [04-core-model.md](./04-core-model.md) §3 と一致。

順序：
```
<styleSheet>
  <numFmts>
  <fonts count="N">
  <fills count="N">
  <borders count="N">
  <cellStyleXfs count="N">
  <cellXfs count="N">
  <cellStyles count="N">
  <dxfs count="N">
  <tableStyles>
  <colors>
</styleSheet>
```

read は schema 経由で plain object を構築し、その後 `_*IdByKey` の dedup index を再構築する。

### 3.1 受け入れ条件

- [ ] 各組み込みスタイルが round-trip 等価
- [ ] 100 個のフォント・100 個のセルスタイルがある xlsx で round-trip
- [ ] write 後の cellXfs プールサイズが open 元と同等（openpyxl は dedup を保証しないため、単調減少のみ確認）

## 4. sharedStrings.xml

read は SAX：
```ts
async function readSharedStrings(stream: ReadableStream<Uint8Array>): Promise<{ table: string[]; index: Map<string, number> }>;
```

`<si>` の中に `<t>`（plain）か `<r>`（rich text run）を含む。rich の場合は `RichText` として保存（`{ kind: 'rich-text', runs }`）。

write は `<sst count uniqueCount>` を一度に出すため、worksheet 書き出し中に蓄積し、最後に flush する。

### 4.1 受け入れ条件

- [ ] empty-string、whitespace-only、特殊 Unicode（emoji、CJK）の round-trip
- [ ] rich-text の round-trip（runs の順序、空 run、空文字列 run）
- [ ] write 中に dedup されている（同じ文字列を 100 万回書いても sst は 1 entry）

## 5. worksheet の read/write

### 5.1 read（通常モード）

参照: openpyxl `worksheet/_reader.py`。

実装は **SAX iterparse**：
```ts
// src/worksheet/reader.ts
export async function parseWorksheet(stream: ReadableStream<Uint8Array>, ctx: ParseContext): Promise<Worksheet>;

interface ParseContext {
  workbook: Workbook;
  sharedStrings: string[];
  rels: Relationships;
  options: LoadOptions;
}
```

SAX イベントに対する dispatcher：
- `<dimension>` → ws の `_dimensionsHint` に保存（最適化用）
- `<sheetView>` → ws.views[]
- `<sheetFormatPr>` → ws.defaultColumnWidth / defaultRowHeight
- `<col>` → columnDimensions
- `<row>` → row 単位で cells を集めて `ws.rows.set(rowIdx, cellsByCol)`
- `<c>` 配下の `<v>` `<f>` `<is>` を組み立てて Cell を生成
- `<mergeCells>` → ws.mergedCells
- `<conditionalFormatting>` → ws.conditionalFormatting（フェーズ5）
- `<dataValidations>` → ws.dataValidations（フェーズ5）
- `<hyperlinks>` → ws.hyperlinks（フェーズ5）
- `<pageSetup>` etc.
- `<drawing r:id="rIdN">` → ws._drawingRel = rels.get('rIdN')（フェーズ6）

cell 値の解釈：
- `t="s"`: shared string index → `sharedStrings[i]`
- `t="b"`: '0'/'1' → boolean
- `t="str"`: plain string formula 結果
- `t="inlineStr"`: `<is>` 配下の rich text or plain
- `t="e"`: error code
- 既定（`t` 省略 or `n`）: number。styleId が date format なら Date に変換（オプションで serial のままも可）

shared formula:
- 初出（`<f t="shared" si="0" ref="A1:A10">`）: 当該 cell には formula を保存し、Translator を `ctx.sharedFormulas[0] = { origin, formula }` にキャッシュ
- 後続（`<f t="shared" si="0"/>` のみ）: `translateFormula(cachedFormula, origin, currentCell)` で展開

### 5.2 write（通常モード）

参照: openpyxl `worksheet/_writer.py`。

`XmlStreamWriter` を介して以下の順で出力：

```
<worksheet>
  <sheetPr/>?
  <dimension ref="A1:Z99"/>
  <sheetViews>
  <sheetFormatPr/>
  <cols>
  <sheetData>
    <row r="1" ...>
      <c r="A1" s="2" t="s"><v>0</v></c>
      ...
    </row>
    ...
  </sheetData>
  <sheetProtection/>?
  <protectedRanges/>?
  <autoFilter/>?
  <mergeCells>
  <conditionalFormatting>*
  <dataValidations>?
  <hyperlinks>?
  <printOptions/>
  <pageMargins/>
  <pageSetup/>
  <headerFooter>?
  <rowBreaks>?
  <colBreaks>?
  <drawing r:id="rId1"/>?
  <legacyDrawing r:id="rId2"/>?
  <legacyDrawingHF r:id=".../>?
  <picture/>?
  <oleObjects/>?
  <controls/>?
  <tableParts>?
  <extLst/>?
</worksheet>
```

各 cell の出力はホットパス：
```ts
// 専用関数（schema 経由しない）
function writeCellXml(out: XmlChunkBuffer, c: Cell, sharedStrings: SharedStringsAccumulator, datesAsSerial: boolean): void;
```

文字列値は **shared strings に追加して index を `<v>` に書く**。空文字列・1 byte 値の極小最適化は不要。

### 5.3 dimension

`<dimension ref="A1:Z99"/>` は **書き出し時に実セル走査して算出**。read 時の値は信用しない（壊れている xlsx も多い）。

### 5.4 受け入れ条件

- [ ] 100 万 cell（1 シート）の write が `~5s`、read が `~10s`（M1）
- [ ] shared formula を含む sheet の round-trip 等価
- [ ] inline string を読み出した後 shared string として書き出すケースが等価（XML 形式は変わるが意味は同じ）
- [ ] 空 row / 空 cell が抜けない
- [ ] `data_only` オプションで formula が cached value に置き換わって保存される

## 6. ハイパーリンク・コメント以外の周辺要素（read-only / 最低限の write）

フェーズ3 ではフル対応は不要。**読み込んで `Worksheet` の plain field に格納** さえできていれば、書き戻しは plain XML passthrough で済ませる：

| 要素 | フェーズ3 | フェーズ5 で完成 |
|------|----------|----------------|
| AutoFilter | passthrough | フル対応 |
| DataValidation | passthrough | フル対応 |
| Hyperlink | passthrough（read 時にセル値の URL 結合のみ） | フル対応 |
| ConditionalFormatting | passthrough | フル対応 |
| Tables | passthrough | フル対応 |
| Comments | passthrough | フル対応 |
| Drawings | passthrough | フェーズ6 |
| Charts | passthrough | フェーズ6 |

**passthrough 実装**: 該当 XML 要素全体を `Uint8Array` で `wb.passthrough` に保存。書き戻し時にそのまま該当位置に挿入。`Worksheet` 上の `_passthroughXml: Map<string, Uint8Array>` に保持する設計。

## 7. defined names / external links / pivot caches（passthrough）

フェーズ3 では **読み書き保持** のみ。編集 API は提供しない（フェーズ5 / 7 で順次拡張）。

## 8. テスト戦略

### 8.1 ゴールデンフィクスチャ

- openpyxl の `tests/data/genuine/*.xlsx` を `tests/fixtures/genuine/` にハードリンク（または symlink）
- 同 `reader/`, `writer/` も
- 自前で生成するフィクスチャは `tests/fixtures/xlsx-kit/` に置く（ファイルプレフィクス `oxjs-`）

### 8.2 round-trip テスト構造

```ts
describe('round-trip: empty-with-styles.xlsx', () => {
  it('reads then writes equivalent XML', async () => {
    const orig = await readZip('tests/fixtures/genuine/empty-with-styles.xlsx');
    const wb = await loadWorkbook(fromBytes(await orig.fullBytes()));
    const out = toBuffer();
    await saveWorkbook(wb, out);
    const written = await readZip(fromBytes(out.result()));
    for (const path of XML_PARTS_TO_COMPARE) {
      const a = canonicalize(orig.read(path));
      const b = canonicalize(written.read(path));
      expect(b).toEqual(a);
    }
  });
});
```

`canonicalize`: 名前空間 prefix 統一、属性順序、空白を正規化（[10-testing.md](./10-testing.md) §3）。

### 8.3 受け入れ条件

- [ ] `tests/fixtures/genuine/` のうち、フェーズ3 範囲（charts/drawings/pivot を除く）の全フィクスチャで round-trip pass
- [ ] read → write → read の二往復で意味的内容（cell value、style index 集合、merged ranges など）が一致
- [ ] 既知の不一致（passthrough 経由したもの）はホワイトリストで明示

## 9. 完了条件（フェーズ3 全体）

- [ ] §2〜§7 の各受け入れ条件
- [ ] フェーズ1, 2 の回帰なし
- [ ] バンドルサイズ予算内
- [ ] LibreOffice / Excel / Google Sheets で開けるかを **手動 QA** で確認（出力 xlsx をコミット時にアーティファクト化）
