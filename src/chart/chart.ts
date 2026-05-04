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

/** Discriminator for chart kinds — only `bar` for stage-1. */
export type ChartKind = BarChart;

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
