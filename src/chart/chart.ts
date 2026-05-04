import type { ShapeProperties } from '../drawing/dml/shape-properties';
import type { TextBody } from '../drawing/dml/text';

// ChartML data model. Per docs/plan/08-charts-drawings.md §5.
//
// **Stage 1**: BarChart end-to-end. The full 17 SpreadsheetML chart
// kinds + 8 chartex kinds land across the upcoming iterations; this
// commit introduces the shared shape so future kinds can plug in
// without breaking the public surface:
//
//   ChartSpace
//     ├─ title?: string
//     ├─ legend?: { position: ... }
//     └─ plotArea
//          ├─ chart: BarChart | LineChart | PieChart | …  (discriminated union)
//          ├─ catAx?: CategoryAxis  (or Date / Series axis)
//          └─ valAx?: ValueAxis
//
// References to worksheet ranges (`Sheet1!$A$1:$A$5`) are kept as
// strings; the optional `cache` field carries the last-known data so
// charts can be rendered without reading the source data.

export type LegendPosition = 'r' | 't' | 'l' | 'b' | 'tr';
export type GroupingType = 'clustered' | 'stacked' | 'percentStacked' | 'standard';
export type BarDirection = 'bar' | 'col';

/** Reference to a worksheet range plus an optional client-side cache of the resolved values. */
export interface NumericRef {
  /** Worksheet-qualified range string, e.g. `Sheet1!$B$1:$B$5`. */
  ref: string;
  /** Optional cached numeric values — Excel writes these for offline rendering. */
  cache?: number[];
  /** Optional `formatCode` Excel uses when rendering each cached value. */
  formatCode?: string;
}

export interface CategoryRef {
  ref: string;
  /** Whether the cache is numeric or string. */
  cacheKind?: 'num' | 'str';
  /** String values (when `cacheKind === 'str'`) or numeric values. */
  cache?: ReadonlyArray<string | number>;
  formatCode?: string;
}

export interface BarSeries {
  /** 0-based slot in the chart (`<c:idx>`). */
  idx: number;
  /** Render order (`<c:order>`). Usually equals `idx`. */
  order: number;
  /** Series title — either a static string or a cell reference. */
  tx?: { kind: 'literal'; value: string } | { kind: 'ref'; ref: string };
  /** Per-series shape properties (fill / line / effects). */
  spPr?: ShapeProperties;
  /** Categories. */
  cat?: CategoryRef;
  /** Values (always required for a bar series). */
  val: NumericRef;
}

export interface BarChart {
  kind: 'bar';
  /** `bar` for horizontal bars, `col` for vertical columns. */
  barDir: BarDirection;
  /** Excel default is `clustered`. */
  grouping: GroupingType;
  varyColors?: boolean;
  series: BarSeries[];
  /** Bar gap width in % of bar width (Excel default 150). */
  gapWidth?: number;
  /** Internal axis ids. The category and value axes carry the same numbers. */
  axIds: [number, number];
}

export interface LineSeries extends BarSeries {
  /** Per-series smoothing toggle. */
  smooth?: boolean;
}

export interface LineChart {
  kind: 'line';
  grouping: GroupingType;
  varyColors?: boolean;
  series: LineSeries[];
  /** Whether to round corners between data points (chart-level default). */
  smooth?: boolean;
  axIds: [number, number];
}

export interface AreaChart {
  kind: 'area';
  grouping: GroupingType;
  varyColors?: boolean;
  series: BarSeries[];
  axIds: [number, number];
}

export interface PieChart {
  kind: 'pie';
  varyColors?: boolean;
  /** Pie / Doughnut have a single ring of slices — but Excel allows multiple series; we mirror that. */
  series: BarSeries[];
}

export interface DoughnutChart {
  kind: 'doughnut';
  varyColors?: boolean;
  series: BarSeries[];
  /** Hole size in % of outer radius (10..90, Excel default 50). */
  holeSize?: number;
  /** First-slice rotation angle in degrees. */
  firstSliceAng?: number;
}

export type ScatterStyle = 'line' | 'lineMarker' | 'marker' | 'none' | 'smooth' | 'smoothMarker';

export interface ScatterSeries {
  idx: number;
  order: number;
  tx?: BarSeries['tx'];
  spPr?: ShapeProperties;
  xVal?: NumericRef;
  yVal: NumericRef;
  smooth?: boolean;
}

export interface ScatterChart {
  kind: 'scatter';
  scatterStyle: ScatterStyle;
  varyColors?: boolean;
  series: ScatterSeries[];
  axIds: [number, number];
}

export type RadarStyle = 'standard' | 'marker' | 'filled';

export interface RadarChart {
  kind: 'radar';
  radarStyle: RadarStyle;
  varyColors?: boolean;
  series: BarSeries[];
  axIds: [number, number];
}

export interface BubbleSeries {
  idx: number;
  order: number;
  tx?: BarSeries['tx'];
  spPr?: ShapeProperties;
  xVal?: NumericRef;
  yVal: NumericRef;
  /** Bubble size — required for a real bubble chart. */
  bubbleSize: NumericRef;
  /** Per-series 3-D toggle. */
  bubble3D?: boolean;
}

export type BubbleSizeRepresents = 'area' | 'w';

export interface BubbleChart {
  kind: 'bubble';
  varyColors?: boolean;
  series: BubbleSeries[];
  bubble3D?: boolean;
  /** Bubble scale 0..300 %. Excel default is 100. */
  bubbleScale?: number;
  showNegBubbles?: boolean;
  sizeRepresents?: BubbleSizeRepresents;
  axIds: [number, number];
}

export interface StockChart {
  kind: 'stock';
  /** Up to 4 series — typically open / high / low / close. */
  series: BarSeries[];
  hiLowLines?: boolean;
  upDownBars?: boolean;
  axIds: [number, number];
}

export interface SurfaceChart {
  kind: 'surface';
  series: BarSeries[];
  /** Wireframe (line-only) when true; smoothed surface fill when false. */
  wireframe?: boolean;
  /** Surfaces use 3 axes: cat + val + ser. */
  axIds: [number, number, number];
}

export type OfPieType = 'bar' | 'pie';
export type SplitType = 'auto' | 'cust' | 'percent' | 'pos' | 'val';

export interface OfPieChart {
  kind: 'ofPie';
  /** `bar` for "Bar of Pie", `pie` for "Pie of Pie". */
  ofPieType: OfPieType;
  varyColors?: boolean;
  series: BarSeries[];
  gapWidth?: number;
  splitType?: SplitType;
  /** Position threshold paired with `splitType='pos'`. */
  splitPos?: number;
  /** Indices of data points moved to the secondary plot when `splitType='cust'`. */
  custSplit?: number[];
  /** Secondary plot size as % of primary (5..200). */
  secondPieSize?: number;
}

// ---- 3-D chart variants ---------------------------------------------------
//
// 3-D charts share most of their attributes with their 2-D counterparts but
// land on different XML tag names (<c:bar3DChart>, etc) and use 3 axes
// (cat / val / ser). We keep them as distinct discriminated-union variants
// so the chart kind stays type-narrowable.

export interface Bar3DChart {
  kind: 'bar3D';
  barDir: BarDirection;
  grouping: GroupingType;
  varyColors?: boolean;
  series: BarSeries[];
  gapWidth?: number;
  /** Bar 3-D adds a `gapDepth` attribute. */
  gapDepth?: number;
  /** Cluster | percentStacked | stacked … plus 'standard' which 2-D doesn't take. */
  shape?: 'cone' | 'coneToMax' | 'box' | 'cylinder' | 'pyramid' | 'pyramidToMax';
  axIds: [number, number, number];
}

export interface Line3DChart {
  kind: 'line3D';
  grouping: GroupingType;
  varyColors?: boolean;
  series: LineSeries[];
  gapDepth?: number;
  axIds: [number, number, number];
}

export interface Pie3DChart {
  kind: 'pie3D';
  varyColors?: boolean;
  series: BarSeries[];
}

export interface Area3DChart {
  kind: 'area3D';
  grouping: GroupingType;
  varyColors?: boolean;
  series: BarSeries[];
  gapDepth?: number;
  axIds: [number, number, number];
}

export interface Surface3DChart {
  kind: 'surface3D';
  series: BarSeries[];
  wireframe?: boolean;
  axIds: [number, number, number];
}

/** Discriminator union of all SpreadsheetML chart kinds modelled so far. */
export type ChartKind =
  | BarChart
  | LineChart
  | AreaChart
  | PieChart
  | DoughnutChart
  | ScatterChart
  | RadarChart
  | BubbleChart
  | StockChart
  | SurfaceChart
  | OfPieChart
  | Bar3DChart
  | Line3DChart
  | Pie3DChart
  | Area3DChart
  | Surface3DChart;

export interface CategoryAxis {
  axId: number;
  /** Crosses partner axis id. */
  crossAx: number;
  position?: 'b' | 't' | 'l' | 'r';
  delete?: boolean;
  /** Axis-line / tick formatting. */
  spPr?: ShapeProperties;
  /** Tick-label text formatting (default text run + paragraph properties). */
  txPr?: TextBody;
}

export interface ValueAxis {
  axId: number;
  crossAx: number;
  position?: 'b' | 't' | 'l' | 'r';
  delete?: boolean;
  majorGridlines?: boolean;
  spPr?: ShapeProperties;
  txPr?: TextBody;
}

export interface PlotArea {
  chart: ChartKind;
  catAx?: CategoryAxis;
  valAx?: ValueAxis;
  /** Plot-area shape properties (background fill, border line). */
  spPr?: ShapeProperties;
}

export interface Legend {
  position: LegendPosition;
  overlay?: boolean;
  spPr?: ShapeProperties;
  txPr?: TextBody;
}

/** Chart title with full DrawingML formatting support. */
export interface ChartTitle {
  /**
   * Plain title text. When set the serializer emits
   * `<c:tx><c:rich><a:p><a:r><a:t>text</a:t></a:r></c:rich></c:tx>`.
   * Mutually exclusive with `tx`.
   */
  text?: string;
  /** Rich text body — overrides `text` when both are present. */
  tx?: TextBody;
  overlay?: boolean;
  spPr?: ShapeProperties;
  txPr?: TextBody;
}

export interface ChartSpace {
  /** Optional chart title. */
  title?: ChartTitle;
  legend?: Legend;
  plotArea: PlotArea;
  /** Honour the formatting hints in cached numeric data when rendering. */
  plotVisOnly?: boolean;
  /** Display blanks as gap, zero, or span — Excel default is `gap`. */
  dispBlanksAs?: 'gap' | 'zero' | 'span';
  /** Chart-space level shape properties (overall frame). */
  spPr?: ShapeProperties;
  /** Chart-space level default text properties. */
  txPr?: TextBody;
}

export function makeBarChart(opts: {
  barDir?: BarDirection;
  grouping?: GroupingType;
  series?: BarSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
  gapWidth?: number;
}): BarChart {
  return {
    kind: 'bar',
    barDir: opts.barDir ?? 'col',
    grouping: opts.grouping ?? 'clustered',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.gapWidth !== undefined ? { gapWidth: opts.gapWidth } : {}),
  };
}

export function makeBarSeries(opts: {
  idx: number;
  order?: number;
  val: NumericRef;
  cat?: CategoryRef;
  tx?: BarSeries['tx'];
}): BarSeries {
  return {
    idx: opts.idx,
    order: opts.order ?? opts.idx,
    val: opts.val,
    ...(opts.cat ? { cat: opts.cat } : {}),
    ...(opts.tx ? { tx: opts.tx } : {}),
  };
}

export function makeChartSpace(opts: {
  plotArea: PlotArea;
  /** Plain string is wrapped in `{ text }`; pass `ChartTitle` for full formatting. */
  title?: string | ChartTitle;
  legend?: Legend;
  plotVisOnly?: boolean;
  dispBlanksAs?: ChartSpace['dispBlanksAs'];
  spPr?: ShapeProperties;
  txPr?: TextBody;
}): ChartSpace {
  const title: ChartTitle | undefined = typeof opts.title === 'string' ? { text: opts.title } : opts.title;
  return {
    plotArea: opts.plotArea,
    ...(title !== undefined ? { title } : {}),
    ...(opts.legend ? { legend: opts.legend } : {}),
    ...(opts.plotVisOnly !== undefined ? { plotVisOnly: opts.plotVisOnly } : {}),
    ...(opts.dispBlanksAs !== undefined ? { dispBlanksAs: opts.dispBlanksAs } : {}),
    ...(opts.spPr ? { spPr: opts.spPr } : {}),
    ...(opts.txPr ? { txPr: opts.txPr } : {}),
  };
}

export function makeLineChart(opts: {
  grouping?: GroupingType;
  series?: LineSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
  smooth?: boolean;
}): LineChart {
  return {
    kind: 'line',
    grouping: opts.grouping ?? 'standard',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.smooth !== undefined ? { smooth: opts.smooth } : {}),
  };
}

export function makeAreaChart(opts: {
  grouping?: GroupingType;
  series?: BarSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
}): AreaChart {
  return {
    kind: 'area',
    grouping: opts.grouping ?? 'standard',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
  };
}

export function makePieChart(opts: { series?: BarSeries[]; varyColors?: boolean }): PieChart {
  return {
    kind: 'pie',
    series: opts.series ?? [],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
  };
}

export function makeDoughnutChart(opts: {
  series?: BarSeries[];
  varyColors?: boolean;
  holeSize?: number;
  firstSliceAng?: number;
}): DoughnutChart {
  return {
    kind: 'doughnut',
    series: opts.series ?? [],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.holeSize !== undefined ? { holeSize: opts.holeSize } : {}),
    ...(opts.firstSliceAng !== undefined ? { firstSliceAng: opts.firstSliceAng } : {}),
  };
}

export function makeScatterChart(opts: {
  scatterStyle?: ScatterStyle;
  series?: ScatterSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
}): ScatterChart {
  return {
    kind: 'scatter',
    scatterStyle: opts.scatterStyle ?? 'lineMarker',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
  };
}

export function makeScatterSeries(opts: {
  idx: number;
  order?: number;
  tx?: BarSeries['tx'];
  xVal?: NumericRef;
  yVal: NumericRef;
  smooth?: boolean;
}): ScatterSeries {
  return {
    idx: opts.idx,
    order: opts.order ?? opts.idx,
    yVal: opts.yVal,
    ...(opts.tx ? { tx: opts.tx } : {}),
    ...(opts.xVal ? { xVal: opts.xVal } : {}),
    ...(opts.smooth !== undefined ? { smooth: opts.smooth } : {}),
  };
}

export function makeRadarChart(opts: {
  radarStyle?: RadarStyle;
  series?: BarSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
}): RadarChart {
  return {
    kind: 'radar',
    radarStyle: opts.radarStyle ?? 'standard',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
  };
}

export function makeBubbleChart(opts: {
  series?: BubbleSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
  bubble3D?: boolean;
  bubbleScale?: number;
  showNegBubbles?: boolean;
  sizeRepresents?: BubbleSizeRepresents;
}): BubbleChart {
  return {
    kind: 'bubble',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.bubble3D !== undefined ? { bubble3D: opts.bubble3D } : {}),
    ...(opts.bubbleScale !== undefined ? { bubbleScale: opts.bubbleScale } : {}),
    ...(opts.showNegBubbles !== undefined ? { showNegBubbles: opts.showNegBubbles } : {}),
    ...(opts.sizeRepresents !== undefined ? { sizeRepresents: opts.sizeRepresents } : {}),
  };
}

export function makeBubbleSeries(opts: {
  idx: number;
  order?: number;
  tx?: BarSeries['tx'];
  xVal?: NumericRef;
  yVal: NumericRef;
  bubbleSize: NumericRef;
  bubble3D?: boolean;
}): BubbleSeries {
  return {
    idx: opts.idx,
    order: opts.order ?? opts.idx,
    yVal: opts.yVal,
    bubbleSize: opts.bubbleSize,
    ...(opts.tx ? { tx: opts.tx } : {}),
    ...(opts.xVal ? { xVal: opts.xVal } : {}),
    ...(opts.bubble3D !== undefined ? { bubble3D: opts.bubble3D } : {}),
  };
}

export function makeStockChart(opts: {
  series?: BarSeries[];
  axIds?: [number, number];
  hiLowLines?: boolean;
  upDownBars?: boolean;
}): StockChart {
  return {
    kind: 'stock',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.hiLowLines !== undefined ? { hiLowLines: opts.hiLowLines } : {}),
    ...(opts.upDownBars !== undefined ? { upDownBars: opts.upDownBars } : {}),
  };
}

export function makeSurfaceChart(opts: {
  series?: BarSeries[];
  wireframe?: boolean;
  axIds?: [number, number, number];
}): SurfaceChart {
  return {
    kind: 'surface',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2, 3],
    ...(opts.wireframe !== undefined ? { wireframe: opts.wireframe } : {}),
  };
}

export function makeOfPieChart(opts: {
  ofPieType?: OfPieType;
  series?: BarSeries[];
  varyColors?: boolean;
  gapWidth?: number;
  splitType?: SplitType;
  splitPos?: number;
  custSplit?: number[];
  secondPieSize?: number;
}): OfPieChart {
  return {
    kind: 'ofPie',
    ofPieType: opts.ofPieType ?? 'pie',
    series: opts.series ?? [],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.gapWidth !== undefined ? { gapWidth: opts.gapWidth } : {}),
    ...(opts.splitType !== undefined ? { splitType: opts.splitType } : {}),
    ...(opts.splitPos !== undefined ? { splitPos: opts.splitPos } : {}),
    ...(opts.custSplit ? { custSplit: opts.custSplit } : {}),
    ...(opts.secondPieSize !== undefined ? { secondPieSize: opts.secondPieSize } : {}),
  };
}

export function makeBar3DChart(opts: {
  barDir?: BarDirection;
  grouping?: GroupingType;
  series?: BarSeries[];
  axIds?: [number, number, number];
  varyColors?: boolean;
  gapWidth?: number;
  gapDepth?: number;
  shape?: Bar3DChart['shape'];
}): Bar3DChart {
  return {
    kind: 'bar3D',
    barDir: opts.barDir ?? 'col',
    grouping: opts.grouping ?? 'clustered',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2, 3],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.gapWidth !== undefined ? { gapWidth: opts.gapWidth } : {}),
    ...(opts.gapDepth !== undefined ? { gapDepth: opts.gapDepth } : {}),
    ...(opts.shape !== undefined ? { shape: opts.shape } : {}),
  };
}

export function makeLine3DChart(opts: {
  grouping?: GroupingType;
  series?: LineSeries[];
  axIds?: [number, number, number];
  varyColors?: boolean;
  gapDepth?: number;
}): Line3DChart {
  return {
    kind: 'line3D',
    grouping: opts.grouping ?? 'standard',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2, 3],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.gapDepth !== undefined ? { gapDepth: opts.gapDepth } : {}),
  };
}

export function makePie3DChart(opts: { series?: BarSeries[]; varyColors?: boolean }): Pie3DChart {
  return {
    kind: 'pie3D',
    series: opts.series ?? [],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
  };
}

export function makeArea3DChart(opts: {
  grouping?: GroupingType;
  series?: BarSeries[];
  axIds?: [number, number, number];
  varyColors?: boolean;
  gapDepth?: number;
}): Area3DChart {
  return {
    kind: 'area3D',
    grouping: opts.grouping ?? 'standard',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2, 3],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.gapDepth !== undefined ? { gapDepth: opts.gapDepth } : {}),
  };
}

export function makeSurface3DChart(opts: {
  series?: BarSeries[];
  wireframe?: boolean;
  axIds?: [number, number, number];
}): Surface3DChart {
  return {
    kind: 'surface3D',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2, 3],
    ...(opts.wireframe !== undefined ? { wireframe: opts.wireframe } : {}),
  };
}
