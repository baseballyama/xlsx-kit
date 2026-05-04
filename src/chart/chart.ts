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

/** Discriminator union of all stage-1 chart kinds. */
export type ChartKind = BarChart | LineChart | AreaChart | PieChart | DoughnutChart | ScatterChart | RadarChart;

export interface CategoryAxis {
  axId: number;
  /** Crosses partner axis id. */
  crossAx: number;
  position?: 'b' | 't' | 'l' | 'r';
  delete?: boolean;
}

export interface ValueAxis {
  axId: number;
  crossAx: number;
  position?: 'b' | 't' | 'l' | 'r';
  delete?: boolean;
  majorGridlines?: boolean;
}

export interface PlotArea {
  chart: ChartKind;
  catAx?: CategoryAxis;
  valAx?: ValueAxis;
}

export interface Legend {
  position: LegendPosition;
  overlay?: boolean;
}

export interface ChartSpace {
  /** Optional chart title (plain string in stage-1). */
  title?: string;
  legend?: Legend;
  plotArea: PlotArea;
  /** Honour the formatting hints in cached numeric data when rendering. */
  plotVisOnly?: boolean;
  /** Display blanks as gap, zero, or span — Excel default is `gap`. */
  dispBlanksAs?: 'gap' | 'zero' | 'span';
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
  title?: string;
  legend?: Legend;
  plotVisOnly?: boolean;
  dispBlanksAs?: ChartSpace['dispBlanksAs'];
}): ChartSpace {
  return {
    plotArea: opts.plotArea,
    ...(opts.title !== undefined ? { title: opts.title } : {}),
    ...(opts.legend ? { legend: opts.legend } : {}),
    ...(opts.plotVisOnly !== undefined ? { plotVisOnly: opts.plotVisOnly } : {}),
    ...(opts.dispBlanksAs !== undefined ? { dispBlanksAs: opts.dispBlanksAs } : {}),
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
