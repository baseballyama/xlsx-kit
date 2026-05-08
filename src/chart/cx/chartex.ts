import type { ShapeProperties } from '../../drawing/dml/shape-properties';
import type { TextBody } from '../../drawing/dml/text';

// Chartex (cx:) data model. Per docs/plan/08-charts-drawings.md §6.
//
// Chartex covers the eight Excel-2016 chart kinds that aren't part of
// ECMA-376: Sunburst, Treemap, Waterfall, Histogram, Pareto, Funnel,
// Box-and-Whisker, and Region Map. Unlike the legacy `c:` chart space
// where each kind has its own `<c:barChart>`/`<c:lineChart>` element,
// chartex puts the discriminator on the series — `<cx:series layoutId="...">` —
// and stores all source data once at the top of the document under
// `<cx:chartData>` so multiple series can share the same numbers.

/** Chartex layout discriminator. Lives on `<cx:series layoutId="...">`. */
export type CxLayoutId =
  | 'clusteredColumn'
  | 'waterfall'
  | 'sunburst'
  | 'treemap'
  | 'boxWhisker'
  | 'pareto'
  | 'paretoLine'
  | 'regionMap'
  | 'funnel';

/** Numeric data point inside a chartex `<cx:numDim>` / `<cx:strDim>` level. */
export interface CxPoint {
  idx: number;
  v: string;
}

/** Numeric dimension (val / colorVal / size / x / y). */
export interface CxNumDim {
  kind: 'num';
  /** Dimension role attribute (`val`, `colorVal`, `size`, `x`, `y`). */
  type: string;
  /** Cell reference, e.g. `Sheet1!$A$1:$A$5`. */
  f?: string;
  /** Optional `dir="col"|"row"` formula direction hint. */
  dir?: 'col' | 'row';
  ptCount?: number;
  pts: CxPoint[];
  formatCode?: string;
}

/** String dimension (cat / colorStr). */
export interface CxStrDim {
  kind: 'str';
  /** Dimension role attribute (`cat`, `colorStr`). */
  type: string;
  f?: string;
  dir?: 'col' | 'row';
  ptCount?: number;
  pts: CxPoint[];
  formatCode?: string;
}

export type CxDim = CxNumDim | CxStrDim;

/** Source data block, referenced by `<cx:dataId val="N"/>` on series. */
export interface CxData {
  id: number;
  dims: CxDim[];
}

/** External data linkage (rare — most charts inline the cache). */
export interface CxExternalData {
  rId: string;
  autoUpdate?: boolean;
}

export interface CxChartData {
  externalData?: CxExternalData;
  data: CxData[];
}

/** Layout-properties union, keyed on the series' layoutId family. */
export type CxLayoutPr =
  | { kind: 'waterfall'; subtotalIdx: number[] }
  | {
      kind: 'binning';
      binCountAuto?: boolean;
      binCount?: number;
      binSize?: number;
      intervalClosed?: 'r' | 'l';
      underflow?: number;
      overflow?: number;
    }
  | { kind: 'parentLabel'; layout: 'overlapping' | 'banner' | 'none' }
  | {
      kind: 'visibility';
      meanLine?: boolean;
      meanMarker?: boolean;
      nonoutliers?: boolean;
      outliers?: boolean;
      quartileMethod?: 'exclusive' | 'inclusive';
    }
  | {
      kind: 'region';
      cultureLanguage?: string;
      cultureRegion?: string;
      projectionType?: 'automatic' | 'mercator' | 'miller' | 'albers';
      regionLabelLayout?: 'none' | 'bestFit' | 'showAll';
    };

export type CxDataLabelPos = 'ctr' | 'b' | 'l' | 'r' | 't' | 'inEnd' | 'outEnd' | 'inBase';

export interface CxDataLabels {
  pos?: CxDataLabelPos;
  visibility?: {
    seriesName?: boolean;
    categoryName?: boolean;
    value?: boolean;
  };
}

export interface CxSeries {
  layoutId: CxLayoutId;
  hidden?: boolean;
  ownerIdx?: number;
  formatIdx?: number;
  axisIds?: number[];
  /** Series text — typically a cell reference. */
  tx?: { f?: string; v?: string };
  dataLabels?: CxDataLabels;
  /** Numeric id pointing into chartData.data. */
  dataId?: number;
  layoutPr?: CxLayoutPr;
  /** Per-series shape properties (fill / line / effects). */
  spPr?: ShapeProperties;
  /** Per-series default text properties. */
  txPr?: TextBody;
}

export interface CxAxis {
  id: number;
  hidden?: boolean;
  /** Continuous (value) axis bounds. */
  valScaling?: { min?: number; max?: number };
  /** Category-axis gap width %. */
  catScalingGapWidth?: number;
  majorGridlines?: boolean;
  title?: CxTitle;
  /** Axis-line / tick formatting. */
  spPr?: ShapeProperties;
  /** Tick-label text formatting. */
  txPr?: TextBody;
}

export interface CxTitle {
  pos?: 't' | 'b' | 'l' | 'r';
  align?: 'ctr' | 'l' | 'r';
  overlay?: boolean;
  /** Plain text (rich runs not preserved at this layer). */
  text?: string;
  /** Cell-reference text source. */
  txDataRef?: string;
  /** Title chrome (frame fill / border). */
  spPr?: ShapeProperties;
  /** Title text formatting. */
  txPr?: TextBody;
}

export interface CxLegend {
  pos?: 't' | 'b' | 'l' | 'r' | 'tr';
  align?: 'ctr' | 'l' | 'r';
  overlay?: boolean;
  /** Legend chrome (frame fill / border). */
  spPr?: ShapeProperties;
  /** Legend text formatting. */
  txPr?: TextBody;
}

export interface CxPlotArea {
  series: CxSeries[];
  axes: CxAxis[];
  /** Plot-surface shape properties (background fill / border line). */
  spPr?: ShapeProperties;
}

export interface CxChart {
  title?: CxTitle;
  plotArea: CxPlotArea;
  legend?: CxLegend;
  plotVisOnly?: boolean;
  dispBlanksAs?: 'span' | 'gap' | 'zero';
}

/** Root model for a chartex `xl/charts/chartN.xml` part. */
export interface CxChartSpace {
  kind: 'cxChartSpace';
  chartData: CxChartData;
  chart: CxChart;
  /** Chart-space level shape properties (overall frame). */
  spPr?: ShapeProperties;
  /** Chart-space level default text properties. */
  txPr?: TextBody;
}

// ---- factories --------------------------------------------------------------

export const makeCxNumDim = (opts: {
  type: string;
  f?: string;
  dir?: 'col' | 'row';
  ptCount?: number;
  pts?: CxPoint[];
  formatCode?: string;
}): CxNumDim => ({
  kind: 'num',
  type: opts.type,
  ...(opts.f !== undefined ? { f: opts.f } : {}),
  ...(opts.dir !== undefined ? { dir: opts.dir } : {}),
  ...(opts.ptCount !== undefined ? { ptCount: opts.ptCount } : {}),
  pts: opts.pts ?? [],
  ...(opts.formatCode !== undefined ? { formatCode: opts.formatCode } : {}),
});

export const makeCxStrDim = (opts: {
  type: string;
  f?: string;
  dir?: 'col' | 'row';
  ptCount?: number;
  pts?: CxPoint[];
  formatCode?: string;
}): CxStrDim => ({
  kind: 'str',
  type: opts.type,
  ...(opts.f !== undefined ? { f: opts.f } : {}),
  ...(opts.dir !== undefined ? { dir: opts.dir } : {}),
  ...(opts.ptCount !== undefined ? { ptCount: opts.ptCount } : {}),
  pts: opts.pts ?? [],
  ...(opts.formatCode !== undefined ? { formatCode: opts.formatCode } : {}),
});

export const makeCxData = (id: number, dims: CxDim[]): CxData => ({ id, dims });

export const makeCxSeries = (opts: {
  layoutId: CxLayoutId;
  dataId?: number;
  hidden?: boolean;
  ownerIdx?: number;
  formatIdx?: number;
  axisIds?: number[];
  tx?: { f?: string; v?: string };
  dataLabels?: CxDataLabels;
  layoutPr?: CxLayoutPr;
  spPr?: ShapeProperties;
  txPr?: TextBody;
}): CxSeries => ({
  layoutId: opts.layoutId,
  ...(opts.hidden !== undefined ? { hidden: opts.hidden } : {}),
  ...(opts.ownerIdx !== undefined ? { ownerIdx: opts.ownerIdx } : {}),
  ...(opts.formatIdx !== undefined ? { formatIdx: opts.formatIdx } : {}),
  ...(opts.axisIds !== undefined ? { axisIds: opts.axisIds } : {}),
  ...(opts.tx ? { tx: opts.tx } : {}),
  ...(opts.dataLabels ? { dataLabels: opts.dataLabels } : {}),
  ...(opts.dataId !== undefined ? { dataId: opts.dataId } : {}),
  ...(opts.layoutPr ? { layoutPr: opts.layoutPr } : {}),
  ...(opts.spPr ? { spPr: opts.spPr } : {}),
  ...(opts.txPr ? { txPr: opts.txPr } : {}),
});

export const makeCxAxis = (opts: {
  id: number;
  hidden?: boolean;
  valScaling?: { min?: number; max?: number };
  catScalingGapWidth?: number;
  majorGridlines?: boolean;
  title?: CxTitle;
  spPr?: ShapeProperties;
  txPr?: TextBody;
}): CxAxis => ({
  id: opts.id,
  ...(opts.hidden !== undefined ? { hidden: opts.hidden } : {}),
  ...(opts.valScaling ? { valScaling: opts.valScaling } : {}),
  ...(opts.catScalingGapWidth !== undefined ? { catScalingGapWidth: opts.catScalingGapWidth } : {}),
  ...(opts.majorGridlines !== undefined ? { majorGridlines: opts.majorGridlines } : {}),
  ...(opts.title ? { title: opts.title } : {}),
  ...(opts.spPr ? { spPr: opts.spPr } : {}),
  ...(opts.txPr ? { txPr: opts.txPr } : {}),
});

export const makeCxChartSpace = (opts: {
  series: CxSeries[];
  data?: CxData[];
  axes?: CxAxis[];
  title?: CxTitle;
  legend?: CxLegend;
  plotVisOnly?: boolean;
  dispBlanksAs?: CxChart['dispBlanksAs'];
  externalData?: CxExternalData;
  plotAreaSpPr?: ShapeProperties;
  spPr?: ShapeProperties;
  txPr?: TextBody;
}): CxChartSpace => ({
  kind: 'cxChartSpace',
  chartData: {
    ...(opts.externalData ? { externalData: opts.externalData } : {}),
    data: opts.data ?? [],
  },
  chart: {
    plotArea: {
      series: opts.series,
      axes: opts.axes ?? [],
      ...(opts.plotAreaSpPr ? { spPr: opts.plotAreaSpPr } : {}),
    },
    ...(opts.title ? { title: opts.title } : {}),
    ...(opts.legend ? { legend: opts.legend } : {}),
    ...(opts.plotVisOnly !== undefined ? { plotVisOnly: opts.plotVisOnly } : {}),
    ...(opts.dispBlanksAs !== undefined ? { dispBlanksAs: opts.dispBlanksAs } : {}),
  },
  ...(opts.spPr ? { spPr: opts.spPr } : {}),
  ...(opts.txPr ? { txPr: opts.txPr } : {}),
});

// ---- per-kind convenience factories ----------------------------------------
//
// Each helper builds a CxChartSpace with one CxData block (id 0) carrying
// the appropriate dimension shapes for the layoutId. Callers populate the
// dim point caches directly when they have data on hand.

export const makeSunburstChart = (opts: { catRef?: string; valRef?: string; valFormatCode?: string }): CxChartSpace =>
  makeCxChartSpace({
    data: [
      makeCxData(0, [
        makeCxStrDim({ type: 'cat', ...(opts.catRef !== undefined ? { f: opts.catRef } : {}) }),
        makeCxNumDim({
          type: 'val',
          ...(opts.valRef !== undefined ? { f: opts.valRef } : {}),
          ...(opts.valFormatCode !== undefined ? { formatCode: opts.valFormatCode } : {}),
        }),
      ]),
    ],
    series: [makeCxSeries({ layoutId: 'sunburst', dataId: 0 })],
  });

export const makeTreemapChart = (opts: {
  catRef?: string;
  valRef?: string;
  parentLabelLayout?: 'overlapping' | 'banner' | 'none';
}): CxChartSpace =>
  makeCxChartSpace({
    data: [
      makeCxData(0, [
        makeCxStrDim({ type: 'cat', ...(opts.catRef !== undefined ? { f: opts.catRef } : {}) }),
        makeCxNumDim({ type: 'val', ...(opts.valRef !== undefined ? { f: opts.valRef } : {}) }),
      ]),
    ],
    series: [
      makeCxSeries({
        layoutId: 'treemap',
        dataId: 0,
        layoutPr: { kind: 'parentLabel', layout: opts.parentLabelLayout ?? 'overlapping' },
      }),
    ],
  });

export const makeWaterfallChart = (opts: { catRef?: string; valRef?: string; subtotalIdx?: number[] }): CxChartSpace =>
  makeCxChartSpace({
    data: [
      makeCxData(0, [
        makeCxStrDim({ type: 'cat', ...(opts.catRef !== undefined ? { f: opts.catRef } : {}) }),
        makeCxNumDim({ type: 'val', ...(opts.valRef !== undefined ? { f: opts.valRef } : {}) }),
      ]),
    ],
    series: [
      makeCxSeries({
        layoutId: 'waterfall',
        dataId: 0,
        axisIds: [0, 1],
        layoutPr: { kind: 'waterfall', subtotalIdx: opts.subtotalIdx ?? [] },
      }),
    ],
    axes: [makeCxAxis({ id: 0 }), makeCxAxis({ id: 1 })],
  });

export const makeHistogramChart = (opts: {
  valRef?: string;
  binCountAuto?: boolean;
  binCount?: number;
  binSize?: number;
  intervalClosed?: 'r' | 'l';
  underflow?: number;
  overflow?: number;
}): CxChartSpace =>
  makeCxChartSpace({
    data: [makeCxData(0, [makeCxNumDim({ type: 'val', ...(opts.valRef !== undefined ? { f: opts.valRef } : {}) })])],
    series: [
      makeCxSeries({
        layoutId: 'clusteredColumn',
        dataId: 0,
        axisIds: [0, 1],
        layoutPr: {
          kind: 'binning',
          ...(opts.binCountAuto !== undefined ? { binCountAuto: opts.binCountAuto } : {}),
          ...(opts.binCount !== undefined ? { binCount: opts.binCount } : {}),
          ...(opts.binSize !== undefined ? { binSize: opts.binSize } : {}),
          ...(opts.intervalClosed !== undefined ? { intervalClosed: opts.intervalClosed } : {}),
          ...(opts.underflow !== undefined ? { underflow: opts.underflow } : {}),
          ...(opts.overflow !== undefined ? { overflow: opts.overflow } : {}),
        },
      }),
    ],
    axes: [makeCxAxis({ id: 0 }), makeCxAxis({ id: 1 })],
  });

export const makeParetoChart = (opts: {
  valRef?: string;
  catRef?: string;
  binCountAuto?: boolean;
  binCount?: number;
  binSize?: number;
}): CxChartSpace => {
  const dims: CxDim[] = [makeCxNumDim({ type: 'val', ...(opts.valRef !== undefined ? { f: opts.valRef } : {}) })];
  if (opts.catRef !== undefined) dims.unshift(makeCxStrDim({ type: 'cat', f: opts.catRef }));
  return makeCxChartSpace({
    data: [makeCxData(0, dims)],
    series: [
      makeCxSeries({
        layoutId: 'clusteredColumn',
        dataId: 0,
        axisIds: [0, 1],
        layoutPr: {
          kind: 'binning',
          ...(opts.binCountAuto !== undefined ? { binCountAuto: opts.binCountAuto } : {}),
          ...(opts.binCount !== undefined ? { binCount: opts.binCount } : {}),
          ...(opts.binSize !== undefined ? { binSize: opts.binSize } : {}),
        },
      }),
      makeCxSeries({ layoutId: 'paretoLine', dataId: 0, ownerIdx: 0, axisIds: [0, 2] }),
    ],
    axes: [makeCxAxis({ id: 0 }), makeCxAxis({ id: 1 }), makeCxAxis({ id: 2 })],
  });
};

export const makeFunnelChart = (opts: { catRef?: string; valRef?: string }): CxChartSpace =>
  makeCxChartSpace({
    data: [
      makeCxData(0, [
        makeCxStrDim({ type: 'cat', ...(opts.catRef !== undefined ? { f: opts.catRef } : {}) }),
        makeCxNumDim({ type: 'val', ...(opts.valRef !== undefined ? { f: opts.valRef } : {}) }),
      ]),
    ],
    series: [makeCxSeries({ layoutId: 'funnel', dataId: 0 })],
  });

export const makeBoxWhiskerChart = (opts: {
  catRef?: string;
  valRef?: string;
  meanLine?: boolean;
  meanMarker?: boolean;
  outliers?: boolean;
  nonoutliers?: boolean;
  quartileMethod?: 'exclusive' | 'inclusive';
}): CxChartSpace =>
  makeCxChartSpace({
    data: [
      makeCxData(0, [
        makeCxStrDim({ type: 'cat', ...(opts.catRef !== undefined ? { f: opts.catRef } : {}) }),
        makeCxNumDim({ type: 'val', ...(opts.valRef !== undefined ? { f: opts.valRef } : {}) }),
      ]),
    ],
    series: [
      makeCxSeries({
        layoutId: 'boxWhisker',
        dataId: 0,
        axisIds: [0, 1],
        layoutPr: {
          kind: 'visibility',
          ...(opts.meanLine !== undefined ? { meanLine: opts.meanLine } : {}),
          ...(opts.meanMarker !== undefined ? { meanMarker: opts.meanMarker } : {}),
          ...(opts.outliers !== undefined ? { outliers: opts.outliers } : {}),
          ...(opts.nonoutliers !== undefined ? { nonoutliers: opts.nonoutliers } : {}),
          ...(opts.quartileMethod !== undefined ? { quartileMethod: opts.quartileMethod } : {}),
        },
      }),
    ],
    axes: [makeCxAxis({ id: 0 }), makeCxAxis({ id: 1 })],
  });

export const makeRegionMapChart = (opts: {
  catRef?: string;
  valRef?: string;
  cultureLanguage?: string;
  cultureRegion?: string;
  projectionType?: 'automatic' | 'mercator' | 'miller' | 'albers';
  regionLabelLayout?: 'none' | 'bestFit' | 'showAll';
}): CxChartSpace =>
  makeCxChartSpace({
    data: [
      makeCxData(0, [
        makeCxStrDim({ type: 'cat', ...(opts.catRef !== undefined ? { f: opts.catRef } : {}) }),
        makeCxNumDim({ type: 'val', ...(opts.valRef !== undefined ? { f: opts.valRef } : {}) }),
      ]),
    ],
    series: [
      makeCxSeries({
        layoutId: 'regionMap',
        dataId: 0,
        layoutPr: {
          kind: 'region',
          ...(opts.cultureLanguage !== undefined ? { cultureLanguage: opts.cultureLanguage } : {}),
          ...(opts.cultureRegion !== undefined ? { cultureRegion: opts.cultureRegion } : {}),
          ...(opts.projectionType !== undefined ? { projectionType: opts.projectionType } : {}),
          ...(opts.regionLabelLayout !== undefined ? { regionLabelLayout: opts.regionLabelLayout } : {}),
        },
      }),
    ],
  });
