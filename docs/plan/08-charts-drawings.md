# 08. フェーズ6: 描画 / 画像 / チャート / Chartsheet

**目的**: DrawingML と ChartML を **フル構造化**。とくに ChartML は **「Excel でできることを ooxml-js でも全部できる」** をゴールとする最重要モジュール。
**期間目安**: 8〜12週間（最重量フェーズ）
**前提**: フェーズ1〜5
**完了条件**: ECMA-376 Part 1 §21（DrawingML） / §17.16（ChartML） のうち SpreadsheetML から到達するすべての要素を構造化し、編集 API を提供する。passthrough は **最終手段としてだけ残す**（未知 extLst 等）。

> **このフェーズだけは「passthrough で逃げない」**。Excel でできるチャート機能は ooxml-js でも作れる、編集できる、再現できるのが目標。

## 1. 全体方針

DrawingML と ChartML は OOXML 内で **最も schema 量が多い** 領域。openpyxl は完全網羅していない箇所もあるが、本プロジェクトでは ECMA-376 仕様 + Excel 365 の実挙動 を一次情報として **完全実装** を狙う。

このフェーズでも **クラスは使わず** plain object + free function（[01-architecture.md](./01-architecture.md) §5）。

### 1.1 何を信頼するか

実装の優先順位：
1. **ECMA-376 Part 1（5th edition）の XSD**：`reference/openpyxl/openpyxl/tests/schemas/` または公式 ECMA-376 を一次情報。
2. **Excel 365 の実挙動**：xlsx を作って Excel に開かせ、保存し直したときの差分を観察。
3. **openpyxl の構造**：参考実装。ただし完全網羅していない領域がある。
4. **LibreOffice Calc の挙動**：互換性の確認用。

差分があれば **Excel 365 を正** とする（openpyxl が拾えていない属性は ooxml-js 側で拾う）。

### 1.2 Schema を厚くする

[03-foundations.md](./03-foundations.md) §4 の Schema は **DrawingML / ChartML の細かい属性** を扱うため、以下を拡張する：

- **Choice element**: ECMA-376 で頻出する `<xsd:choice>`。複数候補のうちどれか 1 つを取る要素（`fillProperties` の `noFill` / `solidFill` / `gradFill` / `blipFill` / `pattFill` / `grpFill` 等）。Schema の `kind: 'choice'` バリアントを追加：
  ```ts
  type ElementDef =
    | …
    | { kind: 'choice'; name: string; variants: Array<{ tag: string; xmlNs?: string; schema: () => Schema<unknown> }>; required?: boolean };
  ```
- **Group element**: 複数要素を同時に内包する `<xsd:group>` の参照。再利用される lnRef / fillRef / effectRef / fontRef / styleMatrixReference 等。
- **MinMax + step**: `<xsd:restriction base="xsd:int" min="0" max="100000">` のような ECMA 制約を Schema 側で表現。
- **List type**: 空白区切り数値配列（dash patterns 等）。

### 1.3 編集 API の二段構え

| 階層 | API | 用途 |
|------|-----|------|
| 高水準 | `makeBarChart({ series, …, style: 'colorful-3', smooth: true })` | 共通パターンを 1 行で |
| 中水準 | `makeBarChart`, `setSeriesFill(s, fill)`, `setAxisScale(a, { min, max })` | 装飾要素を組み合わせる |
| 低水準 | `chart.plotArea.charts[0].series[0].spPr.fill = makeSolidFill(...)` | 任意の構造化編集 |

低水準は `Chart` 系の plain object を直接 mutate する形で提供。中・高水準は free function。

### 1.4 namespace 整理

ChartML 系で使う namespace（[03-foundations.md](./03-foundations.md) §3.3 で網羅する）：

```
http://schemas.openxmlformats.org/drawingml/2006/main          (a)
http://schemas.openxmlformats.org/drawingml/2006/chart         (c)
http://schemas.openxmlformats.org/drawingml/2006/chartDrawing  (cdr)
http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing (xdr)
http://schemas.openxmlformats.org/drawingml/2006/picture       (pic)
http://schemas.microsoft.com/office/drawing/2014/chartex       (cx)   # Sunburst, Treemap, Waterfall, Histogram, Pareto, Funnel, Box-Whisker, Map（拡張チャート）
http://schemas.microsoft.com/office/drawing/2010/chart         (c14)
http://schemas.microsoft.com/office/drawing/2012/chart         (c15)
http://schemas.microsoft.com/office/drawing/2017/03/chart      (c16)  # マーカー color theme 等
```

`cx` namespace は **Excel 2016 で追加された 8 種の新チャート**（Sunburst / Treemap / Waterfall / Histogram / Pareto / Funnel / Box-and-Whisker / Map / Filled Map）に対応する。**これらも本フェーズで対応する**（後述 §6.4）。

## 2. 画像（`src/drawing/image.ts`）

### 2.1 型

```ts
export interface XlsxImage {
  bytes: Uint8Array;
  format: 'png' | 'jpeg' | 'gif' | 'bmp' | 'webp' | 'tiff' | 'svg' | 'emf' | 'wmf';
  width: number;
  height: number;
  /** ZIP 内 path（write 時に確定） */
  path?: string;
  /** rels から逆引きするための rId */
  rId?: string;
}

export function loadImage(bytes: Uint8Array, opts?: { format?: XlsxImage['format'] }): XlsxImage;
export function loadImageFromUrl(url: string): Promise<XlsxImage>;       // ブラウザ用 fetch ヘルパ
export function loadImageFromFile(path: string): Promise<XlsxImage>;     // Node 用
```

### 2.2 寸法検出

`image-size` を **dynamic import**：
```ts
async function detectSize(bytes: Uint8Array): Promise<{ width: number; height: number; format: string }> {
  const { default: sizeOf } = await import('image-size');
  const r = sizeOf(bytes);
  return { width: r.width!, height: r.height!, format: r.type! };
}
```

`emf`/`wmf`/`svg` は image-size でも対応されているが、寸法が無いケースは `XlsxImage.{width,height} = 0` を許容（Excel 側でデフォルト寸法を当てる）。

### 2.3 受け入れ条件

- [ ] PNG/JPEG/GIF/BMP/WebP/TIFF/SVG/EMF/WMF の寸法検出
- [ ] format 不明時に bytes のマジックバイトで再判定
- [ ] DPI 情報があれば pt 換算で reflect

## 3. アンカー / SpreadsheetDrawing（§3〜§4 は前バージョン同じ）

### 3.1 アンカー型

```ts
export type DrawingAnchor =
  | { kind: 'absolute'; pos: Point2D; ext: PositiveSize2D }
  | { kind: 'oneCell';  from: AnchorMarker; ext: PositiveSize2D }
  | { kind: 'twoCell';  from: AnchorMarker; to: AnchorMarker; editAs?: 'twoCell' | 'oneCell' | 'absolute' };

export interface AnchorMarker { col: number; colOff: number; row: number; rowOff: number; }
export interface Point2D { x: number; y: number; }                 // EMU
export interface PositiveSize2D { cx: number; cy: number; }        // EMU
```

helper:
```ts
export function makeTwoCellAnchor(from: string, to: string): DrawingAnchor;  // "A1" -> "C5"
export function emuFromPx(px: number): number;
export function pxFromEmu(emu: number): number;
export function emuFromCm(cm: number): number;
export function emuFromInch(inch: number): number;
```

### 3.2 SpreadsheetDrawing

```ts
export interface DrawingPart {
  anchors: DrawingAnchorEntry[];
  rels: Relationships;
}

export type DrawingAnchorEntry = {
  anchor: DrawingAnchor;
  content:
    | { kind: 'picture'; image: XlsxImage; nvProps: NonVisualPictureProps; spPr?: ShapeProperties }
    | { kind: 'chart'; chart: ChartContainer; nvProps: NonVisualGraphicFrameProps }
    | { kind: 'shape'; shape: ShapeData; nvProps: NonVisualShapeProps; txBody?: TextBody }
    | { kind: 'connector'; connector: ConnectorData; nvProps: NonVisualConnectorProps }
    | { kind: 'group'; nvProps: NonVisualGroupShapeProps; grpSpPr: GroupShapeProperties; children: DrawingAnchorEntry[] };
};
```

`'unknown'` バリアントは持たない（Schema 側で全パターン対応）。仕様外の extLst のみ各構造の `extLst?: ExtensionList` フィールドに格納する。

## 4. DrawingML プリミティブ（**全網羅**）

`src/drawing/` 配下の実装網羅リスト。**openpyxl が拾えていないものも含めて ECMA-376 完全準拠** を目指す。

### 4.1 `colors.ts`

```ts
export type DmlColor =
  | { kind: 'srgb'; value: string /* RRGGBB */ }
  | { kind: 'sysClr'; value: string; lastClr?: string }
  | { kind: 'schemeClr'; value: SchemeColorName }
  | { kind: 'prstClr'; value: PresetColorName }
  | { kind: 'hslClr'; hue: number; sat: number; lum: number }
  | { kind: 'scrgbClr'; r: number; g: number; b: number };

export interface DmlColorWithMods {
  base: DmlColor;
  mods: ColorMod[];
}
export type ColorMod =
  | { kind: 'lumMod'; val: number }
  | { kind: 'lumOff'; val: number }
  | { kind: 'satMod'; val: number }
  | { kind: 'satOff'; val: number }
  | { kind: 'hueMod'; val: number }
  | { kind: 'hueOff'; val: number }
  | { kind: 'tint'; val: number }
  | { kind: 'shade'; val: number }
  | { kind: 'alpha'; val: number }
  | { kind: 'alphaMod'; val: number }
  | { kind: 'alphaOff'; val: number }
  | { kind: 'red' | 'green' | 'blue'; val: number }
  | { kind: 'redMod' | 'greenMod' | 'blueMod' | 'redOff' | 'greenOff' | 'blueOff'; val: number }
  | { kind: 'gray' }
  | { kind: 'comp' }
  | { kind: 'inv' }
  | { kind: 'invGamma' }
  | { kind: 'gamma' };
```

`SchemeColorName` は 17 種（`bg1` `tx1` `bg2` `tx2` `accent1`〜`accent6` `hlink` `folHlink` `phClr` `dk1` `lt1` `dk2` `lt2`）。
`PresetColorName` は 140+ 種（aliceBlue, antiqueWhite, …）。

### 4.2 `fill.ts`

```ts
export type Fill =
  | { kind: 'noFill' }
  | { kind: 'solidFill'; color: DmlColorWithMods }
  | { kind: 'gradFill'; flip?: 'x' | 'y' | 'xy' | 'none'; rotWithShape?: boolean; stops: GradientStop[]; lineDir?: { ang: number; scaled: boolean } | { kind: 'path'; pathType: 'shape' | 'circle' | 'rect'; tileRect?: RelativeRect } }
  | { kind: 'blipFill'; blip: Blip; tile?: TileFill; stretch?: { fillRect?: RelativeRect }; srcRect?: RelativeRect; dpi?: number; rotWithShape?: boolean }
  | { kind: 'pattFill'; preset: PresetPattern; fgClr?: DmlColorWithMods; bgClr?: DmlColorWithMods }
  | { kind: 'grpFill' };

export interface GradientStop { pos: number /* 0-100000 */; color: DmlColorWithMods; }
export interface Blip { embedRId?: string; linkRId?: string; cstate?: 'email' | 'screen' | 'print' | 'hqprint'; effects?: BlipEffect[]; }
export type BlipEffect =
  | { kind: 'biLevel'; thresh: number }
  | { kind: 'blur'; rad: number; grow?: boolean }
  | { kind: 'clrChange'; useA?: boolean; clrFrom: DmlColor; clrTo: DmlColor }
  | { kind: 'clrRepl'; color: DmlColor }
  | { kind: 'duotone'; colors: [DmlColor, DmlColor] }
  | { kind: 'fillOverlay'; blend: 'over' | 'mult' | 'screen' | 'darken' | 'lighten'; fill: Fill }
  | { kind: 'grayscl' }
  | { kind: 'hsl'; hue: number; sat: number; lum: number }
  | { kind: 'lum'; bright?: number; contrast?: number }
  | { kind: 'tint'; hue: number; amt: number }
  | { kind: 'alphaModFix'; amt: number };
```

### 4.3 `line.ts`

```ts
export interface LineProperties {
  w?: number;                        // EMU
  cap?: 'rnd' | 'sq' | 'flat';
  cmpd?: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';
  algn?: 'ctr' | 'in';
  fill?: Fill;
  dash?: PresetDash | { kind: 'custDash'; pattern: number[] };
  join?: 'round' | 'bevel' | { kind: 'miter'; lim?: number };
  headEnd?: LineEnd;
  tailEnd?: LineEnd;
}
export type PresetDash = 'solid' | 'dot' | 'dash' | 'lgDash' | 'dashDot' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot' | 'sysDashDot' | 'sysDashDotDot';
export interface LineEnd { type?: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow'; w?: 'sm' | 'med' | 'lg'; len?: 'sm' | 'med' | 'lg'; }
```

### 4.4 `effect.ts`

```ts
export type Effect =
  | { kind: 'blur'; rad: number; grow?: boolean }
  | { kind: 'fillOverlay'; blend: 'over' | 'mult' | 'screen' | 'darken' | 'lighten'; fill: Fill }
  | { kind: 'glow'; rad: number; color: DmlColorWithMods }
  | { kind: 'innerShdw'; blurRad: number; dist: number; dir: number; color: DmlColorWithMods }
  | { kind: 'outerShdw'; blurRad?: number; dist?: number; dir?: number; sx?: number; sy?: number; kx?: number; ky?: number; algn?: 'ctr' | 'tl' | 't' | 'tr' | 'l' | 'r' | 'bl' | 'b' | 'br'; rotWithShape?: boolean; color: DmlColorWithMods }
  | { kind: 'prstShdw'; prst: 'shdw1' | 'shdw2' | /* … 20種 */; dist: number; dir: number; color: DmlColorWithMods }
  | { kind: 'reflection'; blurRad?: number; stA?: number; stPos?: number; endA?: number; endPos?: number; dist?: number; dir?: number; fadeDir?: number; sx?: number; sy?: number; kx?: number; ky?: number; algn?: string; rotWithShape?: boolean }
  | { kind: 'softEdge'; rad: number };

export interface EffectList {
  list: Effect[];                  // 順序保持
}
export interface EffectContainer {
  type: 'tree' | 'sib';
  effects: Effect[];
}
```

### 4.5 `geometry.ts`

```ts
export interface Transform2D { off?: Point2D; ext?: PositiveSize2D; rot?: number; flipH?: boolean; flipV?: boolean; chOff?: Point2D; chExt?: PositiveSize2D; }
export interface PresetGeometry { prst: PresetShape; avLst?: GuideValue[]; }
export interface CustomGeometry { avLst?: GuideValue[]; gdLst?: ShapeGuide[]; ahLst?: AdjustHandle[]; cxnLst?: ConnectionSite[]; rect?: GuideRect; pathLst: GeometryPath[]; }
export interface ShapeGuide { name: string; fmla: string; }
export interface GeometryPath { w?: number; h?: number; fill?: 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess'; stroke?: boolean; extrusionOk?: boolean; commands: PathCommand[]; }
export type PathCommand =
  | { kind: 'moveTo'; pt: Point2D }
  | { kind: 'lnTo'; pt: Point2D }
  | { kind: 'arcTo'; wR: string; hR: string; stAng: string; swAng: string }
  | { kind: 'quadBezTo'; pts: [Point2D, Point2D] }
  | { kind: 'cubicBezTo'; pts: [Point2D, Point2D, Point2D] }
  | { kind: 'close' };

export type PresetShape = 'rect' | 'roundRect' | 'snip1Rect' | /* …195種を網羅 */;
```

PresetShape は **195 種すべて enum 化**（`reference/openpyxl/openpyxl/drawing/geometry.py:Shape.prst` の Set 値）。

### 4.6 `text.ts`（drawing 文脈の RichText）

```ts
export interface TextBody {
  bodyPr: TextBodyProperties;
  lstStyle?: TextListStyle;
  paragraphs: TextParagraph[];
}
export interface TextBodyProperties {
  rot?: number; spcFirstLastPara?: boolean; vertOverflow?: 'overflow' | 'ellipsis' | 'clip'; horzOverflow?: 'overflow' | 'clip';
  vert?: 'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl';
  wrap?: 'none' | 'square'; lIns?: number; tIns?: number; rIns?: number; bIns?: number; numCol?: number; spcCol?: number; rtlCol?: boolean; fromWordArt?: boolean;
  anchor?: 't' | 'ctr' | 'b' | 'just' | 'dist'; anchorCtr?: boolean;
  forceAA?: boolean; upright?: boolean; compatLnSpc?: boolean;
  prstTxWarp?: { prst: PresetTextShape; avLst?: GuideValue[] };
  scene3d?: Scene3DProperties; sp3d?: Shape3DProperties; flatTx?: { z: number };
}
export interface TextParagraph { pPr?: ParagraphProperties; runs: TextRun[]; endParaRPr?: RunProperties; }
export type TextRun =
  | { kind: 'r'; rPr?: RunProperties; t: string }
  | { kind: 'br'; rPr?: RunProperties }
  | { kind: 'fld'; id: string; type?: string; rPr?: RunProperties; t?: string };

export interface ParagraphProperties { marL?: number; marR?: number; lvl?: number; indent?: number; algn?: 'l' | 'ctr' | 'r' | 'just' | 'justLow' | 'dist' | 'thaiDist'; defTabSz?: number; rtl?: boolean; eaLnBrk?: boolean; fontAlgn?: 'auto' | 't' | 'ctr' | 'base' | 'b'; latinLnBrk?: boolean; hangingPunct?: boolean; lnSpc?: TextSpacing; spcBef?: TextSpacing; spcAft?: TextSpacing; tabLst?: TabStop[]; defRPr?: RunProperties; bullet?: BulletProperties; }
export interface RunProperties { kumimoji?: boolean; lang?: string; altLang?: string; sz?: number; b?: boolean; i?: boolean; u?: TextUnderlineType; strike?: 'noStrike' | 'sngStrike' | 'dblStrike'; kern?: number; cap?: 'none' | 'small' | 'all'; spc?: number; normalizeH?: boolean; baseline?: number; noProof?: boolean; dirty?: boolean; err?: boolean; smtClean?: boolean; smtId?: number; bmk?: string; ln?: LineProperties; fill?: Fill; effectLst?: EffectList; effectDag?: EffectContainer; highlight?: DmlColorWithMods; uLnTx?: 'follow' | LineProperties; uFillTx?: 'follow' | Fill; latin?: TextFont; ea?: TextFont; cs?: TextFont; sym?: TextFont; hlinkClick?: HyperlinkInfo; hlinkMouseOver?: HyperlinkInfo; }
```

→ ECMA-376 の `EG_TextRunFormatting` を完全網羅。

### 4.7 `shape-properties.ts`

```ts
export interface ShapeProperties {
  bwMode?: 'clr' | 'auto' | 'gray' | 'ltGray' | 'invGray' | 'grayWhite' | 'blackGray' | 'blackWhite' | 'black' | 'white' | 'hidden';
  xfrm?: Transform2D;
  geometry?: PresetGeometry | CustomGeometry;
  fill?: Fill;
  ln?: LineProperties;
  effects?: EffectList | EffectContainer;
  scene3d?: Scene3DProperties;
  sp3d?: Shape3DProperties;
  extLst?: ExtensionList;
}
```

### 4.8 受け入れ条件

- [ ] ECMA-376 で定義された preset shape (195 種) すべての round-trip
- [ ] gradient / blip / pattern fill の round-trip
- [ ] custom geometry（pathLst）の round-trip
- [ ] glow / shadow / reflection / softEdge / 3D bevel の round-trip
- [ ] WordArt（`prstTxWarp`）の round-trip

## 5. ChartML — フル実装（**最重要**）

### 5.1 ChartSpace（root）

```ts
export interface ChartSpace {
  date1904?: boolean;
  lang?: string;
  roundedCorners?: boolean;
  style?: number;                            // c14:style (1〜48)
  clrMapOvr?: ColorMappingOverride;
  pivotSource?: PivotSource;
  protection?: ChartProtection;
  chart: ChartContainer;
  spPr?: ShapeProperties;
  txPr?: TextBody;
  externalData?: ExternalData;
  printSettings?: ChartPrintSettings;
  userShapes?: { rId: string };              // chartDrawing
  extLst?: ExtensionList;
}

export interface ChartContainer {
  title?: ChartTitle;
  autoTitleDeleted?: boolean;
  pivotFmts?: PivotFormat[];
  view3D?: View3D;
  floor?: Surface3D;
  sideWall?: Surface3D;
  backWall?: Surface3D;
  plotArea: PlotArea;
  legend?: Legend;
  plotVisOnly?: boolean;
  dispBlanksAs?: 'span' | 'gap' | 'zero';
  showDLblsOverMax?: boolean;
  extLst?: ExtensionList;
}
```

### 5.2 PlotArea & チャート本体

```ts
export interface PlotArea {
  layout?: Layout;
  charts: ChartBody[];                  // 1 plot に複数チャート（combo chart）も可
  axes: Axis[];
  dTable?: DataTable;
  spPr?: ShapeProperties;
}

export type ChartBody =
  | BarChartBody | Bar3DChartBody
  | LineChartBody | Line3DChartBody
  | PieChartBody | Pie3DChartBody | DoughnutChartBody | OfPieChartBody
  | AreaChartBody | Area3DChartBody
  | ScatterChartBody | BubbleChartBody
  | RadarChartBody
  | StockChartBody
  | SurfaceChartBody | Surface3DChartBody
  | CxChartBody;                        // chartex namespace（後述 §6.4）
```

各 chart body には ECMA-376 で定義された **すべての属性** をフィールドとして持つ。例：

```ts
export interface BarChartBody {
  kind: 'barChart';
  barDir: 'bar' | 'col';
  grouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
  varyColors?: boolean;
  series: BarSeries[];
  dLbls?: DataLabelList;
  gapWidth?: number;
  overlap?: number;
  serLines?: ChartLines[];
  axId: [number, number];
}

export interface Bar3DChartBody {
  kind: 'bar3DChart';
  barDir: 'bar' | 'col';
  grouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
  varyColors?: boolean;
  series: BarSeries[];
  dLbls?: DataLabelList;
  gapWidth?: number;
  gapDepth?: number;
  shape?: 'cone' | 'coneToMax' | 'box' | 'cylinder' | 'pyramid' | 'pyramidToMax';
  axId: [number, number, number];
}
```

→ 全チャート種類（**標準 11 種 + 3D 6 種 + chartex 8 種 = 25 種**）について同等に厚い型定義を行う。

### 5.3 Series — 全フィールド

```ts
export interface BarSeries {
  idx: number;
  order: number;
  tx?: SeriesText;
  spPr?: ShapeProperties;
  invertIfNegative?: boolean;
  pictureOptions?: PictureOptions;
  dPt?: DataPoint[];
  dLbls?: DataLabelList;
  trendline?: Trendline[];
  errBars?: ErrorBars;
  cat?: AxisDataSource;
  val: NumericDataSource;
  shape?: 'cone' | 'coneToMax' | 'box' | 'cylinder' | 'pyramid' | 'pyramidToMax';
  extLst?: ExtensionList;
}

export interface ScatterSeries {
  idx: number; order: number;
  tx?: SeriesText;
  spPr?: ShapeProperties;
  marker?: Marker;
  dPt?: DataPoint[];
  dLbls?: DataLabelList;
  trendline?: Trendline[];
  errBars?: ErrorBars;
  xVal?: AxisDataSource;
  yVal?: NumericDataSource;
  smooth?: boolean;
  extLst?: ExtensionList;
}

export interface BubbleSeries {
  idx: number; order: number;
  tx?: SeriesText;
  spPr?: ShapeProperties;
  invertIfNegative?: boolean;
  dPt?: DataPoint[];
  dLbls?: DataLabelList;
  trendline?: Trendline[];
  errBars?: ErrorBars;
  xVal: AxisDataSource;
  yVal: NumericDataSource;
  bubbleSize: NumericDataSource;
  bubble3D?: boolean;
  extLst?: ExtensionList;
}

// LineSeries / PieSeries / RadarSeries / StockSeries / AreaSeries / SurfaceSeries も同様。
```

### 5.4 DataPoint（個別書式）

```ts
export interface DataPoint {
  idx: number;
  invertIfNegative?: boolean;
  marker?: Marker;
  bubble3D?: boolean;
  explosion?: number;          // pie/doughnut の切り出し
  spPr?: ShapeProperties;
  pictureOptions?: PictureOptions;
  extLst?: ExtensionList;
}
```

→ Excel で「特定 1 点だけ赤く塗る」が表現可能。

### 5.5 DataLabel

```ts
export interface DataLabelList {
  dLbl?: DataLabel[];                 // 個別ラベル
  delete?: boolean;
  numFmt?: { formatCode: string; sourceLinked?: boolean };
  spPr?: ShapeProperties;
  txPr?: TextBody;
  dLblPos?: 'bestFit' | 'b' | 'ctr' | 'inBase' | 'inEnd' | 'l' | 'outEnd' | 'r' | 't';
  showLegendKey?: boolean;
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  separator?: string;
  showLeaderLines?: boolean;
  leaderLines?: ChartLines;
  extLst?: ExtensionList;
}
export interface DataLabel { idx: number; tx?: { rich?: TextBody; strRef?: StringReference }; layout?: Layout; /* 上記と同じフィールド… */ }
```

### 5.6 Trendline

```ts
export interface Trendline {
  name?: string;
  spPr?: ShapeProperties;
  trendlineType: 'exp' | 'linear' | 'log' | 'movingAvg' | 'poly' | 'power';
  order?: number;             // 多項式・移動平均の次数/期間
  period?: number;
  forward?: number;
  backward?: number;
  intercept?: number;
  dispRSqr?: boolean;
  dispEq?: boolean;
  trendlineLbl?: TrendlineLabel;
}
export interface TrendlineLabel { layout?: Layout; tx?: { rich?: TextBody; strRef?: StringReference }; numFmt?: { formatCode: string; sourceLinked?: boolean }; spPr?: ShapeProperties; txPr?: TextBody; }
```

### 5.7 ErrorBars

```ts
export interface ErrorBars {
  errDir?: 'x' | 'y';
  errBarType: 'both' | 'minus' | 'plus';
  errValType: 'cust' | 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr';
  noEndCap?: boolean;
  plus?: NumericDataSource | { kind: 'numLit'; literals: number[] };
  minus?: NumericDataSource | { kind: 'numLit'; literals: number[] };
  val?: number;
  spPr?: ShapeProperties;
}
```

### 5.8 Axis — 4 種フル

```ts
export type Axis = CategoryAxis | ValueAxis | DateAxis | SeriesAxis;

export interface AxisCommon {
  axId: number;
  scaling: Scaling;
  delete?: boolean;
  axPos: 'b' | 'l' | 'r' | 't';
  majorGridlines?: ChartLines;
  minorGridlines?: ChartLines;
  title?: ChartTitle;
  numFmt?: { formatCode: string; sourceLinked?: boolean };
  majorTickMark?: 'cross' | 'in' | 'none' | 'out';
  minorTickMark?: 'cross' | 'in' | 'none' | 'out';
  tickLblPos?: 'high' | 'low' | 'nextTo' | 'none';
  spPr?: ShapeProperties;
  txPr?: TextBody;
  crossAx: number;
  crosses?: 'autoZero' | 'max' | 'min';
  crossesAt?: number;
  extLst?: ExtensionList;
}

export interface CategoryAxis extends AxisCommon { kind: 'catAx'; auto?: boolean; lblAlgn?: 'ctr' | 'l' | 'r'; lblOffset?: number; tickLblSkip?: number; tickMarkSkip?: number; noMultiLvlLbl?: boolean; }
export interface ValueAxis    extends AxisCommon { kind: 'valAx'; crossBetween?: 'between' | 'midCat'; majorUnit?: number; minorUnit?: number; dispUnits?: DispUnits; }
export interface DateAxis     extends AxisCommon { kind: 'dateAx'; auto?: boolean; lblOffset?: number; baseTimeUnit?: 'days' | 'months' | 'years'; majorUnit?: number; majorTimeUnit?: 'days' | 'months' | 'years'; minorUnit?: number; minorTimeUnit?: 'days' | 'months' | 'years'; }
export interface SeriesAxis   extends AxisCommon { kind: 'serAx'; tickLblSkip?: number; tickMarkSkip?: number; }

export interface Scaling {
  logBase?: number;
  orientation?: 'maxMin' | 'minMax';
  max?: number;
  min?: number;
}
export interface DispUnits {
  custUnit?: number;
  builtInUnit?: 'hundreds' | 'thousands' | 'tenThousands' | 'hundredThousands' | 'millions' | 'tenMillions' | 'hundredMillions' | 'billions' | 'trillions';
  dispUnitsLbl?: { layout?: Layout; tx?: { rich?: TextBody; strRef?: StringReference }; spPr?: ShapeProperties; txPr?: TextBody; };
}
```

### 5.9 Legend / Title / Layout / Marker / ChartLines / DropLines / HiLowLines / UpDownBars

すべて型定義 + Schema を提供：

```ts
export interface Legend { legendPos?: 'b' | 'tr' | 'l' | 'r' | 't'; legendEntry?: LegendEntry[]; layout?: Layout; overlay?: boolean; spPr?: ShapeProperties; txPr?: TextBody; }
export interface LegendEntry { idx: number; delete?: boolean; txPr?: TextBody; }

export interface ChartTitle { tx?: { rich?: TextBody; strRef?: StringReference }; layout?: Layout; overlay?: boolean; spPr?: ShapeProperties; txPr?: TextBody; }
export interface Layout { manualLayout?: ManualLayout; }
export interface ManualLayout { layoutTarget?: 'inner' | 'outer'; xMode?: 'edge' | 'factor'; yMode?: 'edge' | 'factor'; wMode?: 'edge' | 'factor'; hMode?: 'edge' | 'factor'; x?: number; y?: number; w?: number; h?: number; }

export interface Marker { symbol?: 'circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'picture' | 'plus' | 'square' | 'star' | 'triangle' | 'x' | 'auto'; size?: number; spPr?: ShapeProperties; }

export interface ChartLines { spPr?: ShapeProperties; }
export interface UpDownBars { gapWidth?: number; upBars?: { spPr?: ShapeProperties }; downBars?: { spPr?: ShapeProperties }; }
```

### 5.10 3D 設定

```ts
export interface View3D { rotX?: number; hPercent?: number; rotY?: number; depthPercent?: number; rAngAx?: boolean; perspective?: number; }
export interface Surface3D { thickness?: number; spPr?: ShapeProperties; pictureOptions?: PictureOptions; }
export interface Scene3DProperties { camera: { prst: string; fov?: number; zoom?: number; rot?: { lat: number; lon: number; rev: number } }; lightRig: { rig: string; dir: string; rot?: { lat: number; lon: number; rev: number } }; backdrop?: { anchor?: Point3D; norm?: Vector3D; up?: Vector3D }; }
export interface Shape3DProperties { z?: number; extrusionH?: number; contourW?: number; prstMaterial?: string; bevelT?: { w: number; h: number; prst: string }; bevelB?: { w: number; h: number; prst: string }; extrusionClr?: DmlColorWithMods; contourClr?: DmlColorWithMods; }
```

### 5.11 Reference / DataSource

```ts
export interface NumericDataSource {
  source: { kind: 'numRef'; f: string; cache?: NumberCache } | { kind: 'numLit'; literals: NumberLiteral[]; ptCount?: number; formatCode?: string };
}
export interface AxisDataSource {
  source:
    | { kind: 'strRef'; f: string; cache?: StringCache }
    | { kind: 'numRef'; f: string; cache?: NumberCache; formatCode?: string }
    | { kind: 'multiLvlStrRef'; f: string; cache?: MultiLevelStringCache }
    | { kind: 'strLit'; literals: string[]; ptCount?: number }
    | { kind: 'numLit'; literals: NumberLiteral[]; ptCount?: number; formatCode?: string };
}
export interface NumberCache { formatCode?: string; ptCount?: number; pt: Array<{ idx: number; v: number; formatCode?: string }>; }
export interface StringCache { ptCount?: number; pt: Array<{ idx: number; v: string }>; }
export interface MultiLevelStringCache { ptCount?: number; lvl: Array<Array<{ idx: number; v: string }>>; }
```

### 5.12 編集 API（高水準）

```ts
// src/chart/factory.ts
export interface BarChartFactoryOptions {
  series: Array<{ name?: string; values: string | number[]; categories?: string | string[]; fill?: Fill; trendline?: 'linear' | 'log' | 'exp' | 'poly2' | 'poly3' | 'movingAvg-5' }>;
  barDir?: 'bar' | 'col';
  grouping?: 'clustered' | 'stacked' | 'percentStacked';
  title?: string;
  style?: number;                  // c14:style 1-48 のスタイル番号
  legend?: 'b' | 'l' | 'r' | 't' | 'tr' | 'none';
  axes?: { xTitle?: string; yTitle?: string; xMin?: number; xMax?: number; yMin?: number; yMax?: number; logBase?: number };
  showDataLabels?: boolean;
  threeD?: { view: View3D; shape?: BarChartBody['shape'] };
}

export function makeBarChart(opts: BarChartFactoryOptions): ChartSpace;

// 同様に: makeLineChart / makePieChart / makeScatterChart / makeAreaChart / makeRadarChart /
//          makeBubbleChart / makeDoughnutChart / makeStockChart / makeSurfaceChart /
//          makeSunburstChart / makeTreemapChart / makeWaterfallChart / makeHistogramChart /
//          makeParetoChart / makeFunnelChart / makeBoxWhiskerChart / makeMapChart
```

### 5.13 編集 API（中水準）

```ts
// src/chart/edit.ts
export function setSeriesName(chart: ChartSpace, idx: number, name: string): void;
export function setSeriesValues(chart: ChartSpace, idx: number, ref: string): void;
export function setSeriesCategories(chart: ChartSpace, idx: number, ref: string): void;
export function setSeriesFill(chart: ChartSpace, idx: number, fill: Fill): void;
export function setSeriesLine(chart: ChartSpace, idx: number, line: LineProperties): void;
export function setSeriesMarker(chart: ChartSpace, idx: number, marker: Marker): void;
export function addSeries(chart: ChartSpace, series: BarSeries | LineSeries | /* … */): void;
export function removeSeries(chart: ChartSpace, idx: number): void;
export function addTrendline(chart: ChartSpace, seriesIdx: number, t: Trendline): void;
export function addErrorBars(chart: ChartSpace, seriesIdx: number, e: ErrorBars): void;
export function setAxisTitle(chart: ChartSpace, axisRole: 'x' | 'y' | 'y2' | 'z', title: string | TextBody): void;
export function setAxisScale(chart: ChartSpace, axisRole: 'x' | 'y' | 'y2' | 'z', scaling: Partial<Scaling>): void;
export function setAxisNumFmt(chart: ChartSpace, axisRole: 'x' | 'y' | 'y2' | 'z', formatCode: string): void;
export function setLegendPosition(chart: ChartSpace, pos: 'b' | 'l' | 'r' | 't' | 'tr' | 'none'): void;
export function setChartTitle(chart: ChartSpace, title: string | TextBody): void;
export function setChartStyle(chart: ChartSpace, styleId: number /* 1-48 */): void;
export function setDataLabels(chart: ChartSpace, opts: { seriesIdx?: number; show?: { value?: boolean; category?: boolean; series?: boolean; percent?: boolean; bubbleSize?: boolean }; position?: DataLabelList['dLblPos']; numFmt?: string }): void;
export function setView3D(chart: ChartSpace, view: View3D): void;
```

### 5.14 編集 API（低水準）

`ChartSpace` の plain object を直接 mutate することで、上記の中水準で表現できない操作も可能。ただし型は `Readonly<…>` を **付けない**（mutate を許す）。フィールドは凍結しない。

### 5.15 read/write 経路

参照: openpyxl `chart/reader.py:7`, `chart/_chart.py:125`、ECMA-376 §17.16。

**read**:
1. drawing 部品 (`xl/drawings/drawingN.xml`) に `<graphicFrame>` を発見
2. graphic data uri で chart namespace を判別
3. `chart` rels を辿って `xl/charts/chartN.xml` をロード
4. `ChartSpace` schema で deserialize
5. `cx` namespace（chartex）の場合は `xl/charts/chartExN.xml` を別途読む

**write**:
1. `ChartSpace` を schema → XML（chart 名前空間 + 必要な拡張 ns）
2. drawing 側で `<graphicFrame>` の `r:id` を割り当て
3. content type に `application/vnd.openxmlformats-officedocument.drawingml.chart+xml` を登録
4. chartex の場合は `application/vnd.ms-office.chartex+xml`

### 5.16 受け入れ条件（非常に厚い）

- [ ] **チャート 25 種すべての round-trip**（標準 11 + 3D 6 + chartex 8）
- [ ] 全 axis 種類（catAx / valAx / dateAx / serAx）の round-trip
- [ ] series の `dPt`（点別書式）の round-trip
- [ ] trendline 6 種すべての round-trip
- [ ] error bars 全 type / valType の round-trip
- [ ] dataLabel の dLblPos 9 種すべての round-trip
- [ ] legend 5 種位置すべての round-trip
- [ ] view3D / floor / sideWall / backWall / scene3d / sp3d の round-trip
- [ ] gradient / picture / pattern fill の chart background 適用
- [ ] glow / shadow / soft-edge / 3D bevel の round-trip
- [ ] preset shape 195 種に対する spPr 適用
- [ ] strRef / numRef / multiLvlStrRef / strLit / numLit の round-trip
- [ ] cache（numCache / strCache）の数値・文字列が消えない
- [ ] secondary axis（c2 軸）を伴う combo chart の round-trip（barChart + lineChart on 同 plot）
- [ ] chart style 1〜48 の round-trip
- [ ] **Excel 365 で開いた時、視覚的に同等**（手動 QA で各 chart を Excel 起動・スクショ・LibreOffice でも確認）

## 6. 拡張チャート（chartex namespace）

### 6.1 8 種

ECMA-376 にはなく、Microsoft Office 2016 以降で追加された：

| chartex 種類 | 説明 |
|------------|------|
| `sunburst` | サンバースト |
| `treemap` | ツリーマップ |
| `waterfall` | ウォーターフォール |
| `histogram` / `pareto` | ヒストグラム / パレート |
| `funnel` | ファネル |
| `boxWhisker` | 箱ひげ |
| `clusteredColumn`/`paretoLine`/`stackedColumn`/`region`/`series` etc. | （series の layout マッピングで実現） |
| `regionMap` | マップチャート（地理） |

### 6.2 別途 schema を持つ

`xl/charts/chartExN.xml` は ECMA-376 と完全に異なる schema（`http://schemas.microsoft.com/office/drawing/2014/chartex`）。専用 reader/writer を `src/chart/cx/` に配置。

```ts
// src/chart/cx/chartex.ts
export interface CxChartSpace {
  chartData: { externalData?: { rId: string; autoUpdate: boolean }; data: CxData[] };
  chart: CxChart;
  clrMapOvr?: ColorMappingOverride;
  spPr?: ShapeProperties;
  txPr?: TextBody;
  printSettings?: ChartPrintSettings;
  extLst?: ExtensionList;
}

export interface CxChart {
  title?: ChartTitle;
  plotArea: { plotAreaRegion: { plotSurface?: ShapeProperties; series: CxSeries[]; axis?: CxAxis[] }; layout?: Layout };
  legend?: Legend;
  plotVisOnly?: boolean;
  dispBlanksAs?: 'span' | 'gap' | 'zero';
}

export interface CxSeries {
  layoutId: 'clusteredColumn' | 'waterfall' | 'sunburst' | 'treemap' | 'boxWhisker' | 'pareto' | 'regionMap' | 'funnel';
  hidden?: boolean;
  ownerIdx?: number;
  /* layoutId ごとに可変フィールド（複数 axis、subtotal 等） */
  layoutPr?: CxLayoutProperties;
  axisId?: number[];
  tx?: { txData?: { f?: string; v?: string } };
  dataPt?: CxDataPoint[];
  dataLabels?: CxDataLabel[];
  spPr?: ShapeProperties;
}
```

`cx` 系の編集 API は中水準まで提供（`makeWaterfallChart` 等）。

### 6.3 受け入れ条件

- [ ] chartex 8 種すべての round-trip
- [ ] mixed cluster (chartex + 標準 chart 混在) の round-trip
- [ ] Excel で chartex 図を作って ooxml-js で読んでも seriesData が消えない

## 7. Chartsheet（`src/chartsheet/`）

```ts
export interface Chartsheet {
  title: string;
  views: ChartsheetView[];
  protection?: ChartsheetProtection;
  pageMargins: PageMargins;
  pageSetup: PageSetup;
  headerFooter: HeaderFooter;
  drawing: DrawingPart;          // AbsoluteAnchor で chart を 1 個保持
  drawingHF?: { rId: string };
  picture?: { rId: string };
  smartTags?: { rId: string };
  customSheetViews?: CustomChartsheetView[];
  webPublishItems?: WebPublishItems[];
}
```

編集 API：
```ts
export function addChartsheet(wb: Workbook, title: string, chart: ChartSpace): Chartsheet;
export function setChartsheetView(cs: Chartsheet, view: ChartsheetView): void;
```

### 7.1 受け入れ条件

- [ ] chart-only sheet の round-trip
- [ ] AbsoluteAnchor の位置・サイズが保持される
- [ ] customSheetViews / publishedItems の round-trip

## 8. UserShapes（chartDrawing）

`xl/drawings/chartDrawing*.xml` は chart 上に貼られた **追加図形**（注釈用テキストボックス、矢印など）。`cdr:chartDrawing` namespace。`ChartSpace.userShapes` の rId 経由でリンク。

```ts
export interface ChartDrawing {
  shapes: Array<RelativeAnchor | AbsoluteAnchor>;
}

export type RelativeAnchor = { kind: 'relSizeAnchor'; from: ChartRelativeMarker; to: ChartRelativeMarker; content: ShapeData | TextBody | PictureFrame; }
export type AbsoluteAnchor = { kind: 'absSizeAnchor'; from: ChartRelativeMarker; ext: PositiveSize2D; content: ShapeData | TextBody | PictureFrame; }
export interface ChartRelativeMarker { x: number /* 0-1 */; y: number /* 0-1 */; }
```

編集 API：
```ts
export function addChartShape(chart: ChartSpace, shape: ShapeData, anchor: RelativeAnchor | AbsoluteAnchor): void;
export function addChartTextBox(chart: ChartSpace, text: string | TextBody, anchor: RelativeAnchor | AbsoluteAnchor): void;
```

### 8.1 受け入れ条件

- [ ] chartDrawing 上のテキストボックス・図形・画像の round-trip
- [ ] chart 上に矢印を引く編集が ooxml-js で可能、かつ Excel で問題なく開ける

## 9. VML（legacy comments の anchor）

完全な VML パーサは持たない。**legacy comment の VML drawing を XML フラグメントとして round-trip する**だけ。新規 legacy comment 追加時は最小 VML を生成（フェーズ5 §1 を参照）。

> 注：このフェーズは ChartML を最優先で進める。VML は legacy comment 用に最小限のままでよい。

## 10. テスト

### 10.1 ゴールデンフィクスチャ

- openpyxl `chart/tests/data/`, `drawing/tests/data/` の xlsx 群を fixture に追加
- 自前テスト用に：
  - 各標準 chart 種類（11）の最小例
  - 各 3D variant（6）の最小例
  - 各 chartex 種類（8）の最小例
  - combo chart（bar + line）
  - secondary axis 付き
  - trendline / errorBars / dataLabel 各組み合わせ
  - chartDrawing 追加図形付き

### 10.2 視覚的検証

PR 単位で以下を手動 QA：
1. 各 chart 種類の出力 xlsx を Excel 365 / LibreOffice / Google Sheets / WPS で開く
2. スクリーンショットを `tests/qa/charts/` に保管
3. 視覚回帰差分を画像 diff で機械チェック（pixelmatch）

### 10.3 受け入れ条件（フェーズ全体）

- [ ] §2〜§9 の各受け入れ条件
- [ ] チャートの **全 25 種** が round-trip
- [ ] チャート全装飾要素（trendline, errorBar, marker, dLbl, legend, axis, view3D, fill/line/effect）が round-trip
- [ ] 既存フェーズの回帰なし
- [ ] バンドルサイズ予算: `ooxml-js/chart` は ChartML の量が多いため特例で **≤ 120KB min+gz**（基本サブセット）。`ooxml-js/chart/extended` で chartex を分離して **≤ 60KB min+gz** 追加。`ooxml-js/drawing` ≤ 40KB min+gz
- [ ] LibreOffice + Excel 365 + Google Sheets で各 chart が **視覚的に等価**（QA 画像で照合）
