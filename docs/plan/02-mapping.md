# 02. openpyxl ↔ xlsx-kit モジュール対応表

実装中に「この機能は openpyxl のどこにある？」「これは TS のどのモジュールに置く？」を即引きするための表。

`reference/openpyxl/` 配下のパスを基準に記載する。`tests/` ディレクトリは省略している（テストは [10-testing.md](./10-testing.md) で別途整理）。

## 1. 主要ディレクトリの対応

| openpyxl (Python) | xlsx-kit (TS) | 主な責務 | 詳細 |
|------------------|------------------|----------|------|
| `openpyxl/__init__.py` | `src/index.ts` | 公開エントリ。`loadWorkbook`, `createWorkbook`, `save` を再 export | [01-architecture.md](./01-architecture.md) §6 |
| `openpyxl/_constants.py` | `src/constants.ts` | バージョン、URL、メタ | – |
| `openpyxl/compat/` | `src/compat/` | numbers, singleton, abc | フェーズ1 |
| `openpyxl/xml/` | `src/xml/` | XML parse / serialize、namespace、iterparse | フェーズ1 |
| `openpyxl/packaging/` | `src/packaging/` | manifest, relationships, doc properties | フェーズ1 |
| `openpyxl/descriptors/` | `src/schema/` | Schema 定義 + toTree/fromTree（class 不使用、純データ） | フェーズ1 |
| `openpyxl/utils/` | `src/utils/` | coordinate / datetime / units / inference / escape / exceptions | フェーズ1 |
| `openpyxl/cell/` | `src/cell/` | Cell, MergedCell, RichText | フェーズ2 |
| `openpyxl/styles/` | `src/styles/` | Font / Fill / Border / NamedStyle / Stylesheet | フェーズ2 |
| `openpyxl/formatting/` | `src/formatting/` | 条件付き書式 | フェーズ5 |
| `openpyxl/workbook/` | `src/workbook/` | Workbook, DefinedName, Properties | フェーズ3 |
| `openpyxl/worksheet/` | `src/worksheet/` | Worksheet 一式 | フェーズ3 |
| `openpyxl/reader/` | `src/workbook/reader.ts` ほか | Workbook 読み込み | フェーズ3 |
| `openpyxl/writer/` | `src/workbook/writer.ts` ほか | Workbook 書き出し | フェーズ3 |
| `openpyxl/comments/` | `src/comments/` | 通常コメント / threaded コメント | フェーズ5 |
| `openpyxl/formula/` | `src/formula/` | tokenizer, translator | フェーズ3（読み出し時に必要） |
| `openpyxl/drawing/` | `src/drawing/` | 画像 / shape / anchor | フェーズ6 |
| `openpyxl/chart/` | `src/chart/` | Bar/Line/Pie/Scatter/Area 他 | フェーズ6 |
| `openpyxl/chartsheet/` | `src/chartsheet/` | チャート専用シート | フェーズ6 |
| `openpyxl/pivot/` | `src/pivot/` | ピボット（passthrough 主体） | フェーズ7 |

## 2. ファイル単位の対応（主要）

### 2.1 IO / packaging / xml

| openpyxl ファイル | TS ファイル | 備考 |
|-----------------|-----------|------|
| `xml/__init__.py` | `src/xml/index.ts` | LXML/DEFUSEDXML 切替は不要。fast-xml-parser に統一 |
| `xml/functions.py` | `src/xml/parser.ts` + `src/xml/serializer.ts` + `src/xml/namespaces.ts` | parse/serialize/iterparse/namespace 登録 |
| `xml/constants.py` | `src/xml/namespaces.ts` | OOXML 全名前空間 const 定義 |
| `packaging/manifest.py` | `src/packaging/manifest.ts` | `[Content_Types].xml` |
| `packaging/relationship.py` | `src/packaging/relationships.ts` | rels の dedup と rId 付番 |
| `packaging/core.py` | `src/packaging/core.ts` | `docProps/core.xml` (Dublin Core) |
| `packaging/extended.py` | `src/packaging/extended.ts` | `docProps/app.xml` |
| `packaging/custom.py` | `src/packaging/custom.ts` | `docProps/custom.xml` |
| `packaging/interface.py` | （不要） | Python の ABC ベースのため |
| `packaging/workbook.py` | `src/workbook/package.ts` | `xl/workbook.xml` の root schema |

### 2.2 Workbook / Worksheet

| openpyxl ファイル | TS ファイル | 備考 |
|-----------------|-----------|------|
| `workbook/workbook.py` | `src/workbook/workbook.ts` | data + free function（`createWorkbook`, `addSheet`, `removeSheet`, `getSheetByName`, `setActive` 等） |
| `workbook/_writer.py` | `src/workbook/writer.ts` | `xl/workbook.xml`, root rels, `xl/workbook.xml.rels` |
| `workbook/child.py` | `src/workbook/child.ts` | Child sheet meta |
| `workbook/defined_name.py` | `src/workbook/defined-name.ts` | 名前定義 |
| `workbook/properties.py` | `src/workbook/properties.ts` | calcPr, workbookPr |
| `workbook/protection.py` | `src/workbook/protection.ts` | DocumentSecurity |
| `workbook/views.py` | `src/workbook/views.ts` | BookView |
| `workbook/external_link/external.py` | `src/workbook/external-link/external.ts` | external workbook 参照 |
| `workbook/web.py` | `src/workbook/web.ts` | webPublishing 設定 |
| `workbook/function_group.py` | `src/workbook/function-groups.ts` | functionGroups |
| `workbook/smart_tags.py` | `src/workbook/smart-tags.ts` | スマートタグ |
| `worksheet/worksheet.py` | `src/worksheet/worksheet.ts` | Worksheet データ |
| `worksheet/_reader.py` | `src/worksheet/reader.ts` | SAX iterparse でセルストリーム |
| `worksheet/_writer.py` | `src/worksheet/writer.ts` | XmlStreamWriter で `sheetData` を流す |
| `worksheet/_read_only.py` | `src/streaming/read-only-worksheet.ts` | iterRows() 形式の lazy read |
| `worksheet/_write_only.py` | `src/streaming/write-only-worksheet.ts` | appendRow() 形式の streaming write |
| `worksheet/dimensions.py` | `src/worksheet/dimensions.ts` | RowDimension, ColumnDimension |
| `worksheet/cell_range.py` | `src/worksheet/cell-range.ts` | CellRange + 集合演算（free function） |
| `worksheet/merge.py` | `src/worksheet/merge.ts` | MergedCellRange |
| `worksheet/filters.py` | `src/worksheet/filters.ts` | AutoFilter, SortState |
| `worksheet/datavalidation.py` | `src/worksheet/data-validation.ts` | データ検証 |
| `worksheet/hyperlink.py` | `src/worksheet/hyperlink.ts` | ハイパーリンク |
| `worksheet/header_footer.py` | `src/worksheet/header-footer.ts` | ヘッダ・フッタ（`&P` 等） |
| `worksheet/page.py` | `src/worksheet/page.ts` | PrintPageSetup, PageMargins |
| `worksheet/pagebreak.py` | `src/worksheet/page-break.ts` | RowBreak, ColBreak |
| `worksheet/pivot.py` | `src/worksheet/pivot.ts` | Worksheet 上のピボット参照 |
| `worksheet/properties.py` | `src/worksheet/properties.ts` | sheetPr, sheetFormatPr |
| `worksheet/protection.py` | `src/worksheet/protection.ts` | SheetProtection |
| `worksheet/related.py` | `src/worksheet/related.ts` | rels 管理 |
| `worksheet/table.py` | `src/worksheet/table.ts` | Excel テーブル |
| `worksheet/views.py` | `src/worksheet/views.ts` | SheetView, Pane, Selection |
| `worksheet/controls.py` | `src/worksheet/controls.ts` | フォームコントロール（passthrough） |
| `worksheet/copier.py` | `src/worksheet/copier.ts` | ws のコピー |
| `worksheet/etree_worksheet.py` | （writer.ts に統合） | – |
| `worksheet/formula.py` | `src/worksheet/formula.ts` | 配列・データテーブル・共有フォーミュラ表現 |
| `worksheet/header_footer.py` | 上記参照 | – |
| `worksheet/scenario.py` | `src/worksheet/scenario.ts` | passthrough |
| `worksheet/ole.py` | `src/worksheet/ole.ts` | passthrough |

### 2.3 Cell / Style

| openpyxl ファイル | TS ファイル | 備考 |
|-----------------|-----------|------|
| `cell/cell.py` | `src/cell/cell.ts` | Cell + MergedCell |
| `cell/read_only.py` | `src/streaming/read-only-cell.ts` | ReadOnlyCell |
| `cell/rich_text.py` | `src/cell/rich-text.ts` | CellRichText, TextBlock |
| `cell/text.py` | `src/cell/text.ts` | InlineFont, Text |
| `cell/_writer.py` | `src/cell/writer.ts` | Cell → XML |
| `styles/fonts.py` | `src/styles/fonts.ts` + `src/styles/fonts.schema.ts` | Font 値 + Schema |
| `styles/fills.py` | `src/styles/fills.ts` + `.schema.ts` | PatternFill, GradientFill |
| `styles/borders.py` | `src/styles/borders.ts` + `.schema.ts` | Border, Side |
| `styles/alignment.py` | `src/styles/alignment.ts` + `.schema.ts` | Alignment |
| `styles/protection.py` | `src/styles/protection.ts` + `.schema.ts` | Protection |
| `styles/numbers.py` | `src/styles/numbers.ts` | NumberFormat, BUILTIN_FORMATS |
| `styles/colors.py` | `src/styles/colors.ts` | Color (rgb/theme/indexed/auto + tint) |
| `styles/named_styles.py` | `src/styles/named-styles.ts` | NamedStyle |
| `styles/builtins.py` | `src/styles/builtins.ts` | "Normal", "Good", "Bad", … |
| `styles/cell_style.py` | `src/styles/cell-style.ts` | CellStyle |
| `styles/proxy.py` | （不要） | StyleProxy は提供しない（[01-architecture.md](./01-architecture.md) §8） |
| `styles/stylesheet.py` | `src/styles/stylesheet.ts` | 全プールの管理 + dedup |
| `styles/differential.py` | `src/styles/differential.ts` | DXF（条件付き書式と表で使用） |
| `styles/table.py` | `src/styles/table-style.ts` | TableStyle, TableStyleElement |

### 2.4 数式 / 書式

| openpyxl ファイル | TS ファイル | 備考 |
|-----------------|-----------|------|
| `formula/tokenizer.py` | `src/formula/tokenizer.ts` | Excel 関数式の tokenizer |
| `formula/translate.py` | `src/formula/translator.ts` | shared formula の reference shift |
| `formatting/formatting.py` | `src/formatting/conditional.ts` | ConditionalFormatting + List |
| `formatting/rule.py` | `src/formatting/rule.ts` | Rule, ColorScale, DataBar, IconSet |

### 2.5 描画 / チャート

| openpyxl ファイル | TS ファイル | 備考 |
|-----------------|-----------|------|
| `drawing/spreadsheet_drawing.py` | `src/drawing/spreadsheet-drawing.ts` | drawing root |
| `drawing/image.py` | `src/drawing/image.ts` | Pillow → image-size |
| `drawing/picture.py` | `src/drawing/picture.ts` | PictureFrame |
| `drawing/shape.py` | `src/drawing/shape.ts` | Shape |
| `drawing/connector.py` | `src/drawing/connector.ts` | Connector |
| `drawing/anchor.py` | `src/drawing/anchor.ts` | TwoCellAnchor 他 |
| `drawing/geometry.py` | `src/drawing/geometry.ts` | Point2D, Transform2D 等 |
| `drawing/colors.py` | `src/drawing/colors.ts` | DrawingML カラー |
| `drawing/fill.py` | `src/drawing/fill.ts` | DrawingML 塗り |
| `drawing/line.py` | `src/drawing/line.ts` | LineProperties |
| `drawing/effect.py` | `src/drawing/effect.ts` | Glow / Shadow 等 |
| `drawing/text.py` | `src/drawing/text.ts` | RichText (drawing 文脈) |
| `drawing/properties.py` | `src/drawing/properties.ts` | NonVisualDrawingProps |
| `drawing/relation.py` | `src/drawing/relation.ts` | Drawing 内 rels |
| `drawing/xdr.py` | `src/drawing/xdr.ts` | spreadsheetDrawing 専用 namespace |
| `drawing/ole.py` | `src/drawing/ole.ts` | passthrough |
| `chart/chartspace.py` | `src/chart/chart-space.ts` | ChartSpace, ChartContainer |
| `chart/_chart.py` | `src/chart/chart-base.ts` | ChartBase 共通 |
| `chart/bar_chart.py` | `src/chart/bar-chart.ts` | – |
| `chart/line_chart.py` | `src/chart/line-chart.ts` | – |
| `chart/pie_chart.py` | `src/chart/pie-chart.ts` | – |
| `chart/scatter_chart.py` | `src/chart/scatter-chart.ts` | – |
| `chart/area_chart.py` | `src/chart/area-chart.ts` | – |
| `chart/bubble_chart.py` | `src/chart/bubble-chart.ts` | – |
| `chart/doughnut_chart.py` | `src/chart/doughnut-chart.ts` | – |
| `chart/radar_chart.py` | `src/chart/radar-chart.ts` | – |
| `chart/stock_chart.py` | `src/chart/stock-chart.ts` | フル対応（フェーズ6） |
| `chart/surface_chart.py` | `src/chart/surface-chart.ts` | フル対応（フェーズ6） |
| (Excel 拡張: chartex namespace) | `src/chart/cx/*.ts` | Sunburst/Treemap/Waterfall/Histogram/Pareto/Funnel/BoxWhisker/Map（フェーズ6） |
| (Excel UserShapes) | `src/chart/chart-drawing.ts` | チャート上の追加図形・テキストボックス（フェーズ6） |
| `chart/_3d.py` | `src/chart/3d.ts` | View3D, Surface |
| `chart/axis.py` | `src/chart/axis.ts` | TextAxis, NumericAxis, DateAxis |
| `chart/series.py` | `src/chart/series.ts` | Series |
| `chart/series_factory.py` | `src/chart/series-factory.ts` | Series 生成 helper |
| `chart/data_source.py` | `src/chart/data-source.ts` | NumRef, StrRef, NumDataSource |
| `chart/reference.py` | `src/chart/reference.ts` | Reference (worksheet range) |
| `chart/plotarea.py` | `src/chart/plot-area.ts` | PlotArea |
| `chart/legend.py` | `src/chart/legend.ts` | Legend |
| `chart/title.py` | `src/chart/title.ts` | Title |
| `chart/layout.py` | `src/chart/layout.ts` | Layout |
| `chart/marker.py` | `src/chart/marker.ts` | Marker |
| `chart/label.py` | `src/chart/label.ts` | DataLabel |
| `chart/trendline.py` | `src/chart/trendline.ts` | Trendline |
| `chart/error_bar.py` | `src/chart/error-bar.ts` | ErrorBars |
| `chart/print_settings.py` | `src/chart/print-settings.ts` | – |
| `chart/picture.py` | `src/chart/picture.ts` | チャート内画像 |
| `chart/text.py` | `src/chart/text.ts` | RichText（チャート） |
| `chart/shapes.py` | `src/chart/shapes.ts` | GraphicalProperties |
| `chart/descriptors.py` | `src/chart/descriptors.ts` | NestedGapAmount 等の特殊 schema |
| `chart/reader.py` | `src/chart/reader.ts` | read_chart() |
| `chartsheet/chartsheet.py` | `src/chartsheet/chartsheet.ts` | – |
| `chartsheet/properties.py` | `src/chartsheet/properties.ts` | – |
| `chartsheet/protection.py` | `src/chartsheet/protection.ts` | – |
| `chartsheet/views.py` | `src/chartsheet/views.ts` | – |
| `chartsheet/custom.py` | `src/chartsheet/custom.ts` | – |
| `chartsheet/publish.py` | `src/chartsheet/publish.ts` | – |
| `chartsheet/relation.py` | `src/chartsheet/relation.ts` | – |

### 2.6 ピボット

| openpyxl ファイル | TS ファイル | 備考 |
|-----------------|-----------|------|
| `pivot/cache.py` | `src/pivot/cache.ts` | PivotCacheDefinition（passthrough） |
| `pivot/record.py` | `src/pivot/records.ts` | PivotCacheRecords（passthrough） |
| `pivot/table.py` | `src/pivot/table.ts` | PivotTableDefinition |
| `pivot/fields.py` | `src/pivot/fields.ts` | RowField/ColField/DataField/PageField |
| `pivot/common.py` | `src/pivot/common.ts` | 共通フィールド型 |

### 2.7 コメント

| openpyxl ファイル | TS ファイル | 備考 |
|-----------------|-----------|------|
| `comments/comments.py` | `src/comments/comment.ts` | Comment（legacy） |
| `comments/comment_sheet.py` | `src/comments/comment-sheet.ts` | CommentRecord |
| `comments/author.py` | `src/comments/author.ts` | Author 一覧 |

`xl/threadedComments/*` は openpyxl 側でも完全な構造 round-trip ではないので、まず passthrough → 後にフル対応。

## 3. 同名でないものの注意リスト

| openpyxl 用語 | xlsx-kit 用語 | 理由 |
|------------|------------------|------|
| `Workbook(write_only=True)` | `createWriteOnlyWorkbook(sink)` | 引数で振る舞いが大きく変わるためエントリ分割 |
| `load_workbook(...)` | `loadWorkbook(source, opts)` | camelCase |
| `wb.save(filename)` | `saveWorkbook(wb, sink)` | クラスメソッド削除 |
| `Cell.value = X` | そのまま（plain field） | property assignment は許容 |
| `Cell.font = Font(...)` | `setCellFont(wb, cell, font)` | スタイルプール操作のため free function |
| `StyleProxy` | （提供しない） | tree-shake / immutability 優先 |
| `Worksheet.cell(row, col)` | そのまま `cell(ws, row, col)` または `getCell(ws, row, col)` | – |
| `Worksheet.append(iter)` | `appendRow(ws, values)` | – |
| `Reference(ws, ...)` | `makeReference(ws, ...)` | – |
| `Border.with(left=Side(...))` | spread で `{ ...border, left: side }` | immutable update を直接表現 |

## 4. 削除する／TS で再設計する要素

- **`StyleProxy`**: 削除。free function に置換。
- **`@deprecated` デコレータ**: TS ではコメントの `@deprecated` JSDoc + `tsdocConfig` で警告。ランタイム警告は出さない。
- **`Singleton` メタクラス**: シングルトンの代わりに module-level const + `Object.freeze`。
- **`compat.numbers.NUMERIC_TYPES`**: 不要。`typeof v === 'number'` または `Number.isFinite`。Decimal が必要なら `decimal.js-light` を局所 import。
- **`et_xmlfile.xmlfile`**: 自作 `XmlStreamWriter` で代替（[03-foundations.md](./03-foundations.md) §5）。
- **Cython 高速化対象（`worksheet/_reader.py`, `_writer.py`, `utils/cell.py`）**: TS では純関数化＋ホットループ最適化（[01-architecture.md](./01-architecture.md) §7）。

## 5. 「初手で着手しないモジュール」（ストック）

下記は openpyxl にあっても初期フェーズではポートしない。受け入れる代わりに「passthrough 」または「fallback throw」で扱い、フェーズ7以降で順次対応：

- `worksheet/scenario.py`
- `worksheet/ole.py`
- `worksheet/controls.py`（フォームコントロール）
- `chartsheet/custom.py`, `publish.py` の細部
- `pivot/*`（cache/records）の構造編集
- `workbook/external_link/`（読み書きの保持はする）

> 注: チャート系（Stock/Surface/3D variants/chartex）は **フェーズ6 でフル対応**。passthrough にはしない。
