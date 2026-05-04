// xl/charts/chartN.xml read/write. Per docs/plan/08-charts-drawings.md §5.
//
// Stage-1 covers BarChart end-to-end: parse + serialize with title /
// legend / catAx / valAx / series (cat + val refs + numCache /
// strCache). Other chart kinds slot in alongside as their own
// `<c:lineChart>` / `<c:pieChart>` / etc. parsers.

import {
  parseShapeProperties,
  parseTextBody,
  serializeShapeProperties,
  serializeTextBody,
} from '../drawing/dml/dml-xml';
import type { ShapeProperties } from '../drawing/dml/shape-properties';
import type { TextBody } from '../drawing/dml/text';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { CHART_NS, REL_NS, SHEET_DRAWING_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, findChildren, type XmlNode } from '../xml/tree';
import {
  type Area3DChart,
  type AreaChart,
  type Bar3DChart,
  type BarChart,
  type BarDirection,
  type BarSeries,
  type BubbleChart,
  type BubbleSeries,
  type BubbleSizeRepresents,
  type CategoryAxis,
  type CategoryRef,
  type ChartKind,
  type ChartSpace,
  type ChartTitle,
  type DoughnutChart,
  type GroupingType,
  type Legend,
  type LegendPosition,
  type Line3DChart,
  type LineChart,
  type LineSeries,
  makeArea3DChart,
  makeAreaChart,
  makeBar3DChart,
  makeBarChart,
  makeBarSeries,
  makeBubbleChart,
  makeBubbleSeries,
  makeChartSpace,
  makeDoughnutChart,
  makeLine3DChart,
  makeLineChart,
  makeOfPieChart,
  makePie3DChart,
  makePieChart,
  makeRadarChart,
  makeScatterChart,
  makeScatterSeries,
  makeStockChart,
  makeSurface3DChart,
  makeSurfaceChart,
  type NumericRef,
  type OfPieChart,
  type OfPieType,
  type Pie3DChart,
  type PieChart,
  type PlotArea,
  type RadarChart,
  type RadarStyle,
  type ScatterChart,
  type ScatterSeries,
  type ScatterStyle,
  type SplitType,
  type StockChart,
  type Surface3DChart,
  type SurfaceChart,
  type ValueAxis,
} from './chart';

const CHART_SPACE_TAG = `{${CHART_NS}}chartSpace`;
const CHART_TAG = `{${CHART_NS}}chart`;
const TITLE_TAG = `{${CHART_NS}}title`;
const TX_TAG = `{${CHART_NS}}tx`;
const RICH_TAG = `{${CHART_NS}}rich`;
const PLOT_AREA_TAG = `{${CHART_NS}}plotArea`;
const BAR_CHART_TAG = `{${CHART_NS}}barChart`;
const CAT_AX_TAG = `{${CHART_NS}}catAx`;
const VAL_AX_TAG = `{${CHART_NS}}valAx`;
const SER_TAG = `{${CHART_NS}}ser`;
const IDX_TAG = `{${CHART_NS}}idx`;
const ORDER_TAG = `{${CHART_NS}}order`;
const CAT_TAG = `{${CHART_NS}}cat`;
const VAL_TAG = `{${CHART_NS}}val`;
const NUM_REF_TAG = `{${CHART_NS}}numRef`;
const STR_REF_TAG = `{${CHART_NS}}strRef`;
const NUM_CACHE_TAG = `{${CHART_NS}}numCache`;
const STR_CACHE_TAG = `{${CHART_NS}}strCache`;
const F_TAG = `{${CHART_NS}}f`;
const PT_TAG = `{${CHART_NS}}pt`;
const V_TAG = `{${CHART_NS}}v`;
const FORMAT_CODE_TAG = `{${CHART_NS}}formatCode`;
const BAR_DIR_TAG = `{${CHART_NS}}barDir`;
const GROUPING_TAG = `{${CHART_NS}}grouping`;
const VARY_COLORS_TAG = `{${CHART_NS}}varyColors`;
const GAP_WIDTH_TAG = `{${CHART_NS}}gapWidth`;
const AX_ID_TAG = `{${CHART_NS}}axId`;
const DELETE_TAG = `{${CHART_NS}}delete`;
const LINE_CHART_TAG = `{${CHART_NS}}lineChart`;
const AREA_CHART_TAG = `{${CHART_NS}}areaChart`;
const PIE_CHART_TAG = `{${CHART_NS}}pieChart`;
const DOUGHNUT_CHART_TAG = `{${CHART_NS}}doughnutChart`;
const SCATTER_CHART_TAG = `{${CHART_NS}}scatterChart`;
const RADAR_CHART_TAG = `{${CHART_NS}}radarChart`;
const SMOOTH_TAG = `{${CHART_NS}}smooth`;
const HOLE_SIZE_TAG = `{${CHART_NS}}holeSize`;
const FIRST_SLICE_ANG_TAG = `{${CHART_NS}}firstSliceAng`;
const SCATTER_STYLE_TAG = `{${CHART_NS}}scatterStyle`;
const RADAR_STYLE_TAG = `{${CHART_NS}}radarStyle`;
const X_VAL_TAG = `{${CHART_NS}}xVal`;
const Y_VAL_TAG = `{${CHART_NS}}yVal`;
const BUBBLE_CHART_TAG = `{${CHART_NS}}bubbleChart`;
const STOCK_CHART_TAG = `{${CHART_NS}}stockChart`;
const SURFACE_CHART_TAG = `{${CHART_NS}}surfaceChart`;
const BUBBLE_SIZE_TAG = `{${CHART_NS}}bubbleSize`;
const BUBBLE_3D_TAG = `{${CHART_NS}}bubble3D`;
const BUBBLE_SCALE_TAG = `{${CHART_NS}}bubbleScale`;
const SHOW_NEG_BUBBLES_TAG = `{${CHART_NS}}showNegBubbles`;
const SIZE_REPRESENTS_TAG = `{${CHART_NS}}sizeRepresents`;
const HI_LOW_LINES_TAG = `{${CHART_NS}}hiLowLines`;
const UP_DOWN_BARS_TAG = `{${CHART_NS}}upDownBars`;
const WIREFRAME_TAG = `{${CHART_NS}}wireframe`;
const OF_PIE_CHART_TAG = `{${CHART_NS}}ofPieChart`;
const OF_PIE_TYPE_TAG = `{${CHART_NS}}ofPieType`;
const SPLIT_TYPE_TAG = `{${CHART_NS}}splitType`;
const SPLIT_POS_TAG = `{${CHART_NS}}splitPos`;
const CUST_SPLIT_TAG = `{${CHART_NS}}custSplit`;
const SECOND_PIE_SIZE_TAG = `{${CHART_NS}}secondPieSize`;
const SEC_BLOCK_PT_TAG = `{${CHART_NS}}secondaryPt`;
const BAR3D_CHART_TAG = `{${CHART_NS}}bar3DChart`;
const LINE3D_CHART_TAG = `{${CHART_NS}}line3DChart`;
const PIE3D_CHART_TAG = `{${CHART_NS}}pie3DChart`;
const AREA3D_CHART_TAG = `{${CHART_NS}}area3DChart`;
const SURFACE3D_CHART_TAG = `{${CHART_NS}}surface3DChart`;
const GAP_DEPTH_TAG = `{${CHART_NS}}gapDepth`;
const SHAPE_TAG = `{${CHART_NS}}shape`;
const AX_POS_TAG = `{${CHART_NS}}axPos`;
const CROSS_AX_TAG = `{${CHART_NS}}crossAx`;
const MAJOR_GRIDLINES_TAG = `{${CHART_NS}}majorGridlines`;
const LEGEND_TAG = `{${CHART_NS}}legend`;
const LEGEND_POS_TAG = `{${CHART_NS}}legendPos`;
const PLOT_VIS_ONLY_TAG = `{${CHART_NS}}plotVisOnly`;
const DISP_BLANKS_AS_TAG = `{${CHART_NS}}dispBlanksAs`;
const SP_PR_TAG = `{${CHART_NS}}spPr`;
const TX_PR_TAG = `{${CHART_NS}}txPr`;
const OVERLAY_TAG = `{${CHART_NS}}overlay`;
const A_R_TAG = '{http://schemas.openxmlformats.org/drawingml/2006/main}r';
const A_T_TAG = '{http://schemas.openxmlformats.org/drawingml/2006/main}t';

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const escapeText = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const valAttr = (n: XmlNode | undefined): string | undefined => n?.attrs['val'];
const intVal = (n: XmlNode | undefined): number | undefined => {
  const v = valAttr(n);
  if (v === undefined) return undefined;
  const x = Number.parseInt(v, 10);
  return Number.isInteger(x) ? x : undefined;
};
const boolVal = (n: XmlNode | undefined): boolean | undefined => {
  const v = valAttr(n);
  if (v === undefined) return undefined;
  if (v === '1' || v === 'true') return true;
  if (v === '0' || v === 'false') return false;
  return undefined;
};

const parseRichTextString = (richEl: XmlNode): string | undefined => {
  let out = '';
  for (const p of richEl.children) {
    if (typeof p === 'string') continue;
    for (const r of findChildren(p, A_R_TAG)) {
      const t = findChild(r, A_T_TAG);
      if (t?.text) out += t.text;
    }
  }
  return out.length > 0 ? out : undefined;
};

const parseChartTitle = (titleEl: XmlNode): ChartTitle => {
  const out: ChartTitle = {};
  const tx = findChild(titleEl, TX_TAG);
  if (tx) {
    const rich = findChild(tx, RICH_TAG);
    if (rich) {
      // Parse the rich text body fully so callers see the formatted content.
      // For the convenience plain-text shortcut we also run a simple scrape
      // when the body has only literal runs.
      const body = parseTextBody(rich);
      out.tx = body;
      const flat = parseRichTextString(rich);
      if (flat !== undefined) out.text = flat;
    }
  }
  const overlay = boolVal(findChild(titleEl, OVERLAY_TAG));
  if (overlay !== undefined) out.overlay = overlay;
  const spPrEl = findChild(titleEl, SP_PR_TAG);
  if (spPrEl) out.spPr = parseShapeProperties(spPrEl);
  const txPrEl = findChild(titleEl, TX_PR_TAG);
  if (txPrEl) out.txPr = parseTextBody(txPrEl);
  return out;
};

const parseSpPrSlot = (parent: XmlNode): ShapeProperties | undefined => {
  const el = findChild(parent, SP_PR_TAG);
  return el ? parseShapeProperties(el) : undefined;
};

const parseTxPrSlot = (parent: XmlNode): TextBody | undefined => {
  const el = findChild(parent, TX_PR_TAG);
  return el ? parseTextBody(el) : undefined;
};

const parseNumCache = (cacheEl: XmlNode): { values: number[]; formatCode?: string } => {
  const values: number[] = [];
  const fmt = findChild(cacheEl, FORMAT_CODE_TAG)?.text;
  for (const pt of findChildren(cacheEl, PT_TAG)) {
    const v = findChild(pt, V_TAG)?.text;
    if (v !== undefined) values.push(Number.parseFloat(v));
  }
  return fmt !== undefined ? { values, formatCode: fmt } : { values };
};

const parseStrCache = (cacheEl: XmlNode): string[] => {
  const values: string[] = [];
  for (const pt of findChildren(cacheEl, PT_TAG)) {
    const v = findChild(pt, V_TAG)?.text;
    if (v !== undefined) values.push(v);
  }
  return values;
};

const parseNumericRef = (parent: XmlNode, tag: string): NumericRef | undefined => {
  const wrap = findChild(parent, tag);
  if (!wrap) return undefined;
  const numRef = findChild(wrap, NUM_REF_TAG);
  if (!numRef) return undefined;
  const ref = findChild(numRef, F_TAG)?.text ?? '';
  const cacheEl = findChild(numRef, NUM_CACHE_TAG);
  if (!cacheEl) return { ref };
  const { values, formatCode } = parseNumCache(cacheEl);
  return {
    ref,
    cache: values,
    ...(formatCode !== undefined ? { formatCode } : {}),
  };
};

const parseCategoryRef = (parent: XmlNode): CategoryRef | undefined => {
  const cat = findChild(parent, CAT_TAG);
  if (!cat) return undefined;
  const numRef = findChild(cat, NUM_REF_TAG);
  if (numRef) {
    const ref = findChild(numRef, F_TAG)?.text ?? '';
    const cacheEl = findChild(numRef, NUM_CACHE_TAG);
    if (!cacheEl) return { ref, cacheKind: 'num' };
    const { values, formatCode } = parseNumCache(cacheEl);
    return {
      ref,
      cacheKind: 'num',
      cache: values,
      ...(formatCode !== undefined ? { formatCode } : {}),
    };
  }
  const strRef = findChild(cat, STR_REF_TAG);
  if (strRef) {
    const ref = findChild(strRef, F_TAG)?.text ?? '';
    const cacheEl = findChild(strRef, STR_CACHE_TAG);
    if (!cacheEl) return { ref, cacheKind: 'str' };
    return { ref, cacheKind: 'str', cache: parseStrCache(cacheEl) };
  }
  return undefined;
};

const parseSeries = (serEl: XmlNode): BarSeries | undefined => {
  const idx = intVal(findChild(serEl, IDX_TAG));
  const order = intVal(findChild(serEl, ORDER_TAG));
  if (idx === undefined) return undefined;
  const val = parseNumericRef(serEl, VAL_TAG);
  if (!val) return undefined;
  const opts: Parameters<typeof makeBarSeries>[0] = {
    idx,
    val,
  };
  if (order !== undefined) opts.order = order;
  const cat = parseCategoryRef(serEl);
  if (cat) opts.cat = cat;
  const base = makeBarSeries(opts);
  const spPr = parseSpPrSlot(serEl);
  return spPr ? { ...base, spPr } : base;
};

const parseAxIds = (chartEl: XmlNode): [number, number] => {
  const axIdNodes = findChildren(chartEl, AX_ID_TAG);
  return [intVal(axIdNodes[0]) ?? 1, intVal(axIdNodes[1]) ?? 2];
};

const parseBarSeriesList = (chartEl: XmlNode): BarSeries[] => {
  const series: BarSeries[] = [];
  for (const ser of findChildren(chartEl, SER_TAG)) {
    const s = parseSeries(ser);
    if (s) series.push(s);
  }
  return series;
};

const parseBarChart = (barEl: XmlNode): BarChart => {
  const barDir = (valAttr(findChild(barEl, BAR_DIR_TAG)) ?? 'col') as BarDirection;
  const grouping = (valAttr(findChild(barEl, GROUPING_TAG)) ?? 'clustered') as GroupingType;
  const varyColors = boolVal(findChild(barEl, VARY_COLORS_TAG));
  const gapWidth = intVal(findChild(barEl, GAP_WIDTH_TAG));
  return makeBarChart({
    barDir,
    grouping,
    series: parseBarSeriesList(barEl),
    axIds: parseAxIds(barEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(gapWidth !== undefined ? { gapWidth } : {}),
  });
};

const parseLineSeries = (serEl: XmlNode): LineSeries | undefined => {
  const base = parseSeries(serEl);
  if (!base) return undefined;
  const smooth = boolVal(findChild(serEl, SMOOTH_TAG));
  return smooth !== undefined ? { ...base, smooth } : base;
};

const parseLineChart = (lineEl: XmlNode): LineChart => {
  const grouping = (valAttr(findChild(lineEl, GROUPING_TAG)) ?? 'standard') as GroupingType;
  const varyColors = boolVal(findChild(lineEl, VARY_COLORS_TAG));
  const smooth = boolVal(findChild(lineEl, SMOOTH_TAG));
  const series: LineSeries[] = [];
  for (const ser of findChildren(lineEl, SER_TAG)) {
    const s = parseLineSeries(ser);
    if (s) series.push(s);
  }
  return makeLineChart({
    grouping,
    series,
    axIds: parseAxIds(lineEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(smooth !== undefined ? { smooth } : {}),
  });
};

const parseAreaChart = (areaEl: XmlNode): AreaChart => {
  const grouping = (valAttr(findChild(areaEl, GROUPING_TAG)) ?? 'standard') as GroupingType;
  const varyColors = boolVal(findChild(areaEl, VARY_COLORS_TAG));
  return makeAreaChart({
    grouping,
    series: parseBarSeriesList(areaEl),
    axIds: parseAxIds(areaEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
  });
};

const parsePieChart = (pieEl: XmlNode): PieChart => {
  const varyColors = boolVal(findChild(pieEl, VARY_COLORS_TAG));
  return makePieChart({
    series: parseBarSeriesList(pieEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
  });
};

const parseDoughnutChart = (donutEl: XmlNode): DoughnutChart => {
  const varyColors = boolVal(findChild(donutEl, VARY_COLORS_TAG));
  const holeSize = intVal(findChild(donutEl, HOLE_SIZE_TAG));
  const firstSliceAng = intVal(findChild(donutEl, FIRST_SLICE_ANG_TAG));
  return makeDoughnutChart({
    series: parseBarSeriesList(donutEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(holeSize !== undefined ? { holeSize } : {}),
    ...(firstSliceAng !== undefined ? { firstSliceAng } : {}),
  });
};

const parseScatterSeries = (serEl: XmlNode): ScatterSeries | undefined => {
  const idx = intVal(findChild(serEl, IDX_TAG));
  const order = intVal(findChild(serEl, ORDER_TAG));
  if (idx === undefined) return undefined;
  const yVal = parseNumericRef(serEl, Y_VAL_TAG);
  if (!yVal) return undefined;
  const xVal = parseNumericRef(serEl, X_VAL_TAG);
  const smooth = boolVal(findChild(serEl, SMOOTH_TAG));
  const opts: Parameters<typeof makeScatterSeries>[0] = { idx, yVal };
  if (order !== undefined) opts.order = order;
  if (xVal) opts.xVal = xVal;
  if (smooth !== undefined) opts.smooth = smooth;
  const base = makeScatterSeries(opts);
  const spPr = parseSpPrSlot(serEl);
  return spPr ? { ...base, spPr } : base;
};

const parseScatterChart = (scatterEl: XmlNode): ScatterChart => {
  const scatterStyle = (valAttr(findChild(scatterEl, SCATTER_STYLE_TAG)) ?? 'lineMarker') as ScatterStyle;
  const varyColors = boolVal(findChild(scatterEl, VARY_COLORS_TAG));
  const series: ScatterSeries[] = [];
  for (const ser of findChildren(scatterEl, SER_TAG)) {
    const s = parseScatterSeries(ser);
    if (s) series.push(s);
  }
  return makeScatterChart({
    scatterStyle,
    series,
    axIds: parseAxIds(scatterEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
  });
};

const parseRadarChart = (radarEl: XmlNode): RadarChart => {
  const radarStyle = (valAttr(findChild(radarEl, RADAR_STYLE_TAG)) ?? 'standard') as RadarStyle;
  const varyColors = boolVal(findChild(radarEl, VARY_COLORS_TAG));
  return makeRadarChart({
    radarStyle,
    series: parseBarSeriesList(radarEl),
    axIds: parseAxIds(radarEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
  });
};

const parseBubbleSeries = (serEl: XmlNode): BubbleSeries | undefined => {
  const idx = intVal(findChild(serEl, IDX_TAG));
  const order = intVal(findChild(serEl, ORDER_TAG));
  if (idx === undefined) return undefined;
  const yVal = parseNumericRef(serEl, Y_VAL_TAG);
  const bubbleSize = parseNumericRef(serEl, BUBBLE_SIZE_TAG);
  if (!yVal || !bubbleSize) return undefined;
  const xVal = parseNumericRef(serEl, X_VAL_TAG);
  const bubble3D = boolVal(findChild(serEl, BUBBLE_3D_TAG));
  const opts: Parameters<typeof makeBubbleSeries>[0] = { idx, yVal, bubbleSize };
  if (order !== undefined) opts.order = order;
  if (xVal) opts.xVal = xVal;
  if (bubble3D !== undefined) opts.bubble3D = bubble3D;
  const base = makeBubbleSeries(opts);
  const spPr = parseSpPrSlot(serEl);
  return spPr ? { ...base, spPr } : base;
};

const parseBubbleChart = (bubbleEl: XmlNode): BubbleChart => {
  const varyColors = boolVal(findChild(bubbleEl, VARY_COLORS_TAG));
  const bubble3D = boolVal(findChild(bubbleEl, BUBBLE_3D_TAG));
  const bubbleScale = intVal(findChild(bubbleEl, BUBBLE_SCALE_TAG));
  const showNegBubbles = boolVal(findChild(bubbleEl, SHOW_NEG_BUBBLES_TAG));
  const sizeRepresentsRaw = valAttr(findChild(bubbleEl, SIZE_REPRESENTS_TAG));
  const sizeRepresents: BubbleSizeRepresents | undefined =
    sizeRepresentsRaw === 'area' || sizeRepresentsRaw === 'w' ? sizeRepresentsRaw : undefined;
  const series: BubbleSeries[] = [];
  for (const ser of findChildren(bubbleEl, SER_TAG)) {
    const s = parseBubbleSeries(ser);
    if (s) series.push(s);
  }
  return makeBubbleChart({
    series,
    axIds: parseAxIds(bubbleEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(bubble3D !== undefined ? { bubble3D } : {}),
    ...(bubbleScale !== undefined ? { bubbleScale } : {}),
    ...(showNegBubbles !== undefined ? { showNegBubbles } : {}),
    ...(sizeRepresents !== undefined ? { sizeRepresents } : {}),
  });
};

const parseStockChart = (stockEl: XmlNode): StockChart => {
  // hiLowLines / upDownBars are presence flags — Excel emits them with no
  // attributes. We treat the element's existence as the signal.
  const hiLowLines = findChild(stockEl, HI_LOW_LINES_TAG) !== undefined ? true : undefined;
  const upDownBars = findChild(stockEl, UP_DOWN_BARS_TAG) !== undefined ? true : undefined;
  return makeStockChart({
    series: parseBarSeriesList(stockEl),
    axIds: parseAxIds(stockEl),
    ...(hiLowLines !== undefined ? { hiLowLines } : {}),
    ...(upDownBars !== undefined ? { upDownBars } : {}),
  });
};

const parseSurfaceChart = (surfaceEl: XmlNode): SurfaceChart => {
  const wireframe = boolVal(findChild(surfaceEl, WIREFRAME_TAG));
  const axIds3 = parseAxIds3(surfaceEl);
  return makeSurfaceChart({
    series: parseBarSeriesList(surfaceEl),
    axIds: axIds3,
    ...(wireframe !== undefined ? { wireframe } : {}),
  });
};

const parseAxIds3 = (chartEl: XmlNode): [number, number, number] => {
  const axIdNodes = findChildren(chartEl, AX_ID_TAG);
  return [intVal(axIdNodes[0]) ?? 1, intVal(axIdNodes[1]) ?? 2, intVal(axIdNodes[2]) ?? 3];
};

const SPLIT_TYPES: ReadonlyArray<SplitType> = ['auto', 'cust', 'percent', 'pos', 'val'];

const parseOfPieChart = (ofPieEl: XmlNode): OfPieChart => {
  const ofPieType = (valAttr(findChild(ofPieEl, OF_PIE_TYPE_TAG)) ?? 'pie') as OfPieType;
  const varyColors = boolVal(findChild(ofPieEl, VARY_COLORS_TAG));
  const gapWidth = intVal(findChild(ofPieEl, GAP_WIDTH_TAG));
  const splitTypeRaw = valAttr(findChild(ofPieEl, SPLIT_TYPE_TAG));
  const splitType =
    splitTypeRaw && (SPLIT_TYPES as ReadonlyArray<string>).includes(splitTypeRaw)
      ? (splitTypeRaw as SplitType)
      : undefined;
  const splitPos = intVal(findChild(ofPieEl, SPLIT_POS_TAG));
  const secondPieSize = intVal(findChild(ofPieEl, SECOND_PIE_SIZE_TAG));
  // <c:custSplit> contains <c:secondaryPt idx=N/> entries.
  let custSplit: number[] | undefined;
  const custSplitEl = findChild(ofPieEl, CUST_SPLIT_TAG);
  if (custSplitEl) {
    const points: number[] = [];
    for (const pt of findChildren(custSplitEl, SEC_BLOCK_PT_TAG)) {
      const idx = Number.parseInt(pt.attrs['idx'] ?? '', 10);
      if (Number.isInteger(idx)) points.push(idx);
    }
    custSplit = points;
  }
  return makeOfPieChart({
    ofPieType,
    series: parseBarSeriesList(ofPieEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(gapWidth !== undefined ? { gapWidth } : {}),
    ...(splitType !== undefined ? { splitType } : {}),
    ...(splitPos !== undefined ? { splitPos } : {}),
    ...(custSplit !== undefined ? { custSplit } : {}),
    ...(secondPieSize !== undefined ? { secondPieSize } : {}),
  });
};

// 3-D parsers — reuse 2-D helpers where the shape matches.

const parseBar3DChart = (barEl: XmlNode): Bar3DChart => {
  const barDir = (valAttr(findChild(barEl, BAR_DIR_TAG)) ?? 'col') as BarDirection;
  const grouping = (valAttr(findChild(barEl, GROUPING_TAG)) ?? 'clustered') as GroupingType;
  const varyColors = boolVal(findChild(barEl, VARY_COLORS_TAG));
  const gapWidth = intVal(findChild(barEl, GAP_WIDTH_TAG));
  const gapDepth = intVal(findChild(barEl, GAP_DEPTH_TAG));
  const shape = valAttr(findChild(barEl, SHAPE_TAG)) as Bar3DChart['shape'] | undefined;
  return makeBar3DChart({
    barDir,
    grouping,
    series: parseBarSeriesList(barEl),
    axIds: parseAxIds3(barEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(gapWidth !== undefined ? { gapWidth } : {}),
    ...(gapDepth !== undefined ? { gapDepth } : {}),
    ...(shape !== undefined ? { shape } : {}),
  });
};

const parseLine3DChart = (lineEl: XmlNode): Line3DChart => {
  const grouping = (valAttr(findChild(lineEl, GROUPING_TAG)) ?? 'standard') as GroupingType;
  const varyColors = boolVal(findChild(lineEl, VARY_COLORS_TAG));
  const gapDepth = intVal(findChild(lineEl, GAP_DEPTH_TAG));
  const series: LineSeries[] = [];
  for (const ser of findChildren(lineEl, SER_TAG)) {
    const s = parseLineSeries(ser);
    if (s) series.push(s);
  }
  return makeLine3DChart({
    grouping,
    series,
    axIds: parseAxIds3(lineEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(gapDepth !== undefined ? { gapDepth } : {}),
  });
};

const parsePie3DChart = (pieEl: XmlNode): Pie3DChart => {
  const varyColors = boolVal(findChild(pieEl, VARY_COLORS_TAG));
  return makePie3DChart({
    series: parseBarSeriesList(pieEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
  });
};

const parseArea3DChart = (areaEl: XmlNode): Area3DChart => {
  const grouping = (valAttr(findChild(areaEl, GROUPING_TAG)) ?? 'standard') as GroupingType;
  const varyColors = boolVal(findChild(areaEl, VARY_COLORS_TAG));
  const gapDepth = intVal(findChild(areaEl, GAP_DEPTH_TAG));
  return makeArea3DChart({
    grouping,
    series: parseBarSeriesList(areaEl),
    axIds: parseAxIds3(areaEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(gapDepth !== undefined ? { gapDepth } : {}),
  });
};

const parseSurface3DChart = (surfaceEl: XmlNode): Surface3DChart => {
  const wireframe = boolVal(findChild(surfaceEl, WIREFRAME_TAG));
  return makeSurface3DChart({
    series: parseBarSeriesList(surfaceEl),
    axIds: parseAxIds3(surfaceEl),
    ...(wireframe !== undefined ? { wireframe } : {}),
  });
};

const parsePlotChart = (plotAreaEl: XmlNode): ChartKind => {
  const bar = findChild(plotAreaEl, BAR_CHART_TAG);
  if (bar) return parseBarChart(bar);
  const line = findChild(plotAreaEl, LINE_CHART_TAG);
  if (line) return parseLineChart(line);
  const area = findChild(plotAreaEl, AREA_CHART_TAG);
  if (area) return parseAreaChart(area);
  const pie = findChild(plotAreaEl, PIE_CHART_TAG);
  if (pie) return parsePieChart(pie);
  const donut = findChild(plotAreaEl, DOUGHNUT_CHART_TAG);
  if (donut) return parseDoughnutChart(donut);
  const scatter = findChild(plotAreaEl, SCATTER_CHART_TAG);
  if (scatter) return parseScatterChart(scatter);
  const radar = findChild(plotAreaEl, RADAR_CHART_TAG);
  if (radar) return parseRadarChart(radar);
  const bubble = findChild(plotAreaEl, BUBBLE_CHART_TAG);
  if (bubble) return parseBubbleChart(bubble);
  const stock = findChild(plotAreaEl, STOCK_CHART_TAG);
  if (stock) return parseStockChart(stock);
  const surface = findChild(plotAreaEl, SURFACE_CHART_TAG);
  if (surface) return parseSurfaceChart(surface);
  const ofPie = findChild(plotAreaEl, OF_PIE_CHART_TAG);
  if (ofPie) return parseOfPieChart(ofPie);
  const bar3D = findChild(plotAreaEl, BAR3D_CHART_TAG);
  if (bar3D) return parseBar3DChart(bar3D);
  const line3D = findChild(plotAreaEl, LINE3D_CHART_TAG);
  if (line3D) return parseLine3DChart(line3D);
  const pie3D = findChild(plotAreaEl, PIE3D_CHART_TAG);
  if (pie3D) return parsePie3DChart(pie3D);
  const area3D = findChild(plotAreaEl, AREA3D_CHART_TAG);
  if (area3D) return parseArea3DChart(area3D);
  const surface3D = findChild(plotAreaEl, SURFACE3D_CHART_TAG);
  if (surface3D) return parseSurface3DChart(surface3D);
  throw new OpenXmlSchemaError('parseChartXml: no supported chart kind found inside <plotArea>');
};

const parseAxis = (
  axEl: XmlNode,
): {
  axId: number;
  crossAx: number;
  position?: 'b' | 't' | 'l' | 'r';
  delete?: boolean;
  majorGridlines?: boolean;
  spPr?: ShapeProperties;
  txPr?: TextBody;
} => {
  const axId = intVal(findChild(axEl, AX_ID_TAG)) ?? 0;
  const crossAx = intVal(findChild(axEl, CROSS_AX_TAG)) ?? 0;
  const positionRaw = valAttr(findChild(axEl, AX_POS_TAG));
  const validPos = positionRaw === 'b' || positionRaw === 't' || positionRaw === 'l' || positionRaw === 'r';
  const del = boolVal(findChild(axEl, DELETE_TAG));
  const majorGridlines = findChild(axEl, MAJOR_GRIDLINES_TAG) !== undefined ? true : undefined;
  const spPr = parseSpPrSlot(axEl);
  const txPr = parseTxPrSlot(axEl);
  return {
    axId,
    crossAx,
    ...(validPos ? { position: positionRaw as 'b' | 't' | 'l' | 'r' } : {}),
    ...(del !== undefined ? { delete: del } : {}),
    ...(majorGridlines !== undefined ? { majorGridlines } : {}),
    ...(spPr ? { spPr } : {}),
    ...(txPr ? { txPr } : {}),
  };
};

/** Parse a `xl/charts/chartN.xml` payload. */
export function parseChartXml(bytes: Uint8Array | string): ChartSpace {
  const root = parseXml(bytes);
  if (root.name !== CHART_SPACE_TAG) {
    throw new OpenXmlSchemaError(`parseChartXml: root is "${root.name}", expected chartSpace`);
  }
  const chartEl = findChild(root, CHART_TAG);
  if (!chartEl) throw new OpenXmlSchemaError('parseChartXml: <chartSpace> missing <chart>');
  const plotAreaEl = findChild(chartEl, PLOT_AREA_TAG);
  if (!plotAreaEl) throw new OpenXmlSchemaError('parseChartXml: <chart> missing <plotArea>');
  const chart = parsePlotChart(plotAreaEl);
  const catAxEl = findChild(plotAreaEl, CAT_AX_TAG);
  const valAxEl = findChild(plotAreaEl, VAL_AX_TAG);
  const plotAreaSpPr = parseSpPrSlot(plotAreaEl);
  const plotArea: PlotArea = {
    chart,
    ...(catAxEl ? { catAx: parseAxis(catAxEl) as CategoryAxis } : {}),
    ...(valAxEl ? { valAx: parseAxis(valAxEl) as ValueAxis } : {}),
    ...(plotAreaSpPr ? { spPr: plotAreaSpPr } : {}),
  };
  const titleEl = findChild(chartEl, TITLE_TAG);
  const title = titleEl ? parseChartTitle(titleEl) : undefined;
  const legendEl = findChild(chartEl, LEGEND_TAG);
  let legend: Legend | undefined;
  if (legendEl) {
    const posRaw = valAttr(findChild(legendEl, LEGEND_POS_TAG)) as LegendPosition | undefined;
    const overlay = boolVal(findChild(legendEl, OVERLAY_TAG));
    const legendSpPr = parseSpPrSlot(legendEl);
    const legendTxPr = parseTxPrSlot(legendEl);
    legend = {
      position: posRaw ?? 'r',
      ...(overlay !== undefined ? { overlay } : {}),
      ...(legendSpPr ? { spPr: legendSpPr } : {}),
      ...(legendTxPr ? { txPr: legendTxPr } : {}),
    };
  }
  const plotVisOnly = boolVal(findChild(chartEl, PLOT_VIS_ONLY_TAG));
  const dispBlanksAs = valAttr(findChild(chartEl, DISP_BLANKS_AS_TAG)) as ChartSpace['dispBlanksAs'];
  // Top-level spPr / txPr live on chartSpace (sibling of <c:chart>), not inside <c:chart>.
  const spaceSpPr = parseSpPrSlot(root);
  const spaceTxPr = parseTxPrSlot(root);
  return makeChartSpace({
    plotArea,
    ...(title !== undefined ? { title } : {}),
    ...(legend ? { legend } : {}),
    ...(plotVisOnly !== undefined ? { plotVisOnly } : {}),
    ...(dispBlanksAs ? { dispBlanksAs } : {}),
    ...(spaceSpPr ? { spPr: spaceSpPr } : {}),
    ...(spaceTxPr ? { txPr: spaceTxPr } : {}),
  });
}

const serializeNumCache = (cache: ReadonlyArray<number>, formatCode?: string): string => {
  const parts: string[] = ['<c:numCache>'];
  if (formatCode) parts.push(`<c:formatCode>${escapeText(formatCode)}</c:formatCode>`);
  parts.push(`<c:ptCount val="${cache.length}"/>`);
  for (let i = 0; i < cache.length; i++) {
    const v = cache[i] as number;
    parts.push(`<c:pt idx="${i}"><c:v>${escapeText(String(v))}</c:v></c:pt>`);
  }
  parts.push('</c:numCache>');
  return parts.join('');
};

const serializeStrCache = (cache: ReadonlyArray<string | number>): string => {
  const parts: string[] = ['<c:strCache>', `<c:ptCount val="${cache.length}"/>`];
  for (let i = 0; i < cache.length; i++) {
    const v = cache[i] as string | number;
    parts.push(`<c:pt idx="${i}"><c:v>${escapeText(String(v))}</c:v></c:pt>`);
  }
  parts.push('</c:strCache>');
  return parts.join('');
};

const serializeNumericRef = (tag: string, ref: NumericRef): string => {
  const inner = ref.cache !== undefined ? serializeNumCache(ref.cache, ref.formatCode) : '';
  return `<c:${tag}><c:numRef><c:f>${escapeText(ref.ref)}</c:f>${inner}</c:numRef></c:${tag}>`;
};

const serializeCategoryRef = (cat: CategoryRef): string => {
  if (
    cat.cacheKind === 'str' ||
    (cat.cacheKind !== 'num' && cat.cache && cat.cache.some((v) => typeof v === 'string'))
  ) {
    const inner = cat.cache !== undefined ? serializeStrCache(cat.cache) : '';
    return `<c:cat><c:strRef><c:f>${escapeText(cat.ref)}</c:f>${inner}</c:strRef></c:cat>`;
  }
  const numericCache = cat.cache as number[] | undefined;
  const inner = numericCache !== undefined ? serializeNumCache(numericCache, cat.formatCode) : '';
  return `<c:cat><c:numRef><c:f>${escapeText(cat.ref)}</c:f>${inner}</c:numRef></c:cat>`;
};

const serializeSeries = (s: BarSeries): string => {
  const parts: string[] = ['<c:ser>', `<c:idx val="${s.idx}"/>`, `<c:order val="${s.order}"/>`];
  if (s.tx) {
    if (s.tx.kind === 'literal') {
      parts.push(`<c:tx><c:strRef><c:f></c:f>${serializeStrCache([s.tx.value])}</c:strRef></c:tx>`);
    } else {
      parts.push(`<c:tx><c:strRef><c:f>${escapeText(s.tx.ref)}</c:f></c:strRef></c:tx>`);
    }
  }
  if (s.spPr) parts.push(serializeShapeProperties(s.spPr));
  if (s.cat) parts.push(serializeCategoryRef(s.cat));
  parts.push(serializeNumericRef('val', s.val));
  parts.push('</c:ser>');
  return parts.join('');
};

const serializeBarChart = (chart: BarChart): string => {
  const parts: string[] = [
    '<c:barChart>',
    `<c:barDir val="${chart.barDir}"/>`,
    `<c:grouping val="${chart.grouping}"/>`,
  ];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  if (chart.gapWidth !== undefined) parts.push(`<c:gapWidth val="${chart.gapWidth}"/>`);
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:barChart>');
  return parts.join('');
};

const serializeLineSeries = (s: LineSeries): string => {
  const base = serializeSeries(s);
  if (s.smooth === undefined) return base;
  // Inject smooth before </c:ser> (cheap inline patch).
  return base.replace('</c:ser>', `<c:smooth val="${s.smooth ? '1' : '0'}"/></c:ser>`);
};

const serializeLineChart = (chart: LineChart): string => {
  const parts: string[] = ['<c:lineChart>', `<c:grouping val="${chart.grouping}"/>`];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeLineSeries(s));
  if (chart.smooth !== undefined) parts.push(`<c:smooth val="${chart.smooth ? '1' : '0'}"/>`);
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:lineChart>');
  return parts.join('');
};

const serializeAreaChart = (chart: AreaChart): string => {
  const parts: string[] = ['<c:areaChart>', `<c:grouping val="${chart.grouping}"/>`];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:areaChart>');
  return parts.join('');
};

const serializePieChart = (chart: PieChart): string => {
  const parts: string[] = ['<c:pieChart>'];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push('</c:pieChart>');
  return parts.join('');
};

const serializeDoughnutChart = (chart: DoughnutChart): string => {
  const parts: string[] = ['<c:doughnutChart>'];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  if (chart.firstSliceAng !== undefined) parts.push(`<c:firstSliceAng val="${chart.firstSliceAng}"/>`);
  if (chart.holeSize !== undefined) parts.push(`<c:holeSize val="${chart.holeSize}"/>`);
  parts.push('</c:doughnutChart>');
  return parts.join('');
};

const serializeScatterSeries = (s: ScatterSeries): string => {
  const parts: string[] = ['<c:ser>', `<c:idx val="${s.idx}"/>`, `<c:order val="${s.order}"/>`];
  if (s.tx) {
    if (s.tx.kind === 'literal') {
      parts.push(`<c:tx><c:strRef><c:f></c:f>${serializeStrCache([s.tx.value])}</c:strRef></c:tx>`);
    } else {
      parts.push(`<c:tx><c:strRef><c:f>${escapeText(s.tx.ref)}</c:f></c:strRef></c:tx>`);
    }
  }
  if (s.spPr) parts.push(serializeShapeProperties(s.spPr));
  if (s.xVal) parts.push(serializeNumericRef('xVal', s.xVal));
  parts.push(serializeNumericRef('yVal', s.yVal));
  if (s.smooth !== undefined) parts.push(`<c:smooth val="${s.smooth ? '1' : '0'}"/>`);
  parts.push('</c:ser>');
  return parts.join('');
};

const serializeScatterChart = (chart: ScatterChart): string => {
  const parts: string[] = ['<c:scatterChart>', `<c:scatterStyle val="${chart.scatterStyle}"/>`];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeScatterSeries(s));
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:scatterChart>');
  return parts.join('');
};

const serializeRadarChart = (chart: RadarChart): string => {
  const parts: string[] = ['<c:radarChart>', `<c:radarStyle val="${chart.radarStyle}"/>`];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:radarChart>');
  return parts.join('');
};

const serializeBubbleSeries = (s: BubbleSeries): string => {
  const parts: string[] = ['<c:ser>', `<c:idx val="${s.idx}"/>`, `<c:order val="${s.order}"/>`];
  if (s.tx) {
    if (s.tx.kind === 'literal') {
      parts.push(`<c:tx><c:strRef><c:f></c:f>${serializeStrCache([s.tx.value])}</c:strRef></c:tx>`);
    } else {
      parts.push(`<c:tx><c:strRef><c:f>${escapeText(s.tx.ref)}</c:f></c:strRef></c:tx>`);
    }
  }
  if (s.spPr) parts.push(serializeShapeProperties(s.spPr));
  if (s.xVal) parts.push(serializeNumericRef('xVal', s.xVal));
  parts.push(serializeNumericRef('yVal', s.yVal));
  parts.push(serializeNumericRef('bubbleSize', s.bubbleSize));
  if (s.bubble3D !== undefined) parts.push(`<c:bubble3D val="${s.bubble3D ? '1' : '0'}"/>`);
  parts.push('</c:ser>');
  return parts.join('');
};

const serializeBubbleChart = (chart: BubbleChart): string => {
  const parts: string[] = ['<c:bubbleChart>'];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeBubbleSeries(s));
  if (chart.bubble3D !== undefined) parts.push(`<c:bubble3D val="${chart.bubble3D ? '1' : '0'}"/>`);
  if (chart.bubbleScale !== undefined) parts.push(`<c:bubbleScale val="${chart.bubbleScale}"/>`);
  if (chart.showNegBubbles !== undefined) parts.push(`<c:showNegBubbles val="${chart.showNegBubbles ? '1' : '0'}"/>`);
  if (chart.sizeRepresents !== undefined) parts.push(`<c:sizeRepresents val="${chart.sizeRepresents}"/>`);
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:bubbleChart>');
  return parts.join('');
};

const serializeStockChart = (chart: StockChart): string => {
  const parts: string[] = ['<c:stockChart>'];
  for (const s of chart.series) parts.push(serializeSeries(s));
  if (chart.hiLowLines) parts.push('<c:hiLowLines/>');
  if (chart.upDownBars) parts.push('<c:upDownBars><c:gapWidth val="150"/></c:upDownBars>');
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:stockChart>');
  return parts.join('');
};

const serializeSurfaceChart = (chart: SurfaceChart): string => {
  const parts: string[] = ['<c:surfaceChart>'];
  if (chart.wireframe !== undefined) parts.push(`<c:wireframe val="${chart.wireframe ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[2]}"/>`);
  parts.push('</c:surfaceChart>');
  return parts.join('');
};

const serializeOfPieChart = (chart: OfPieChart): string => {
  const parts: string[] = ['<c:ofPieChart>', `<c:ofPieType val="${chart.ofPieType}"/>`];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  if (chart.gapWidth !== undefined) parts.push(`<c:gapWidth val="${chart.gapWidth}"/>`);
  if (chart.splitType !== undefined) parts.push(`<c:splitType val="${chart.splitType}"/>`);
  if (chart.splitPos !== undefined) parts.push(`<c:splitPos val="${chart.splitPos}"/>`);
  if (chart.custSplit !== undefined) {
    parts.push('<c:custSplit>');
    for (const idx of chart.custSplit) parts.push(`<c:secondaryPt idx="${idx}"/>`);
    parts.push('</c:custSplit>');
  }
  if (chart.secondPieSize !== undefined) parts.push(`<c:secondPieSize val="${chart.secondPieSize}"/>`);
  parts.push('</c:ofPieChart>');
  return parts.join('');
};

const serializeBar3DChart = (chart: Bar3DChart): string => {
  const parts: string[] = [
    '<c:bar3DChart>',
    `<c:barDir val="${chart.barDir}"/>`,
    `<c:grouping val="${chart.grouping}"/>`,
  ];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  if (chart.gapWidth !== undefined) parts.push(`<c:gapWidth val="${chart.gapWidth}"/>`);
  if (chart.gapDepth !== undefined) parts.push(`<c:gapDepth val="${chart.gapDepth}"/>`);
  if (chart.shape !== undefined) parts.push(`<c:shape val="${chart.shape}"/>`);
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[2]}"/>`);
  parts.push('</c:bar3DChart>');
  return parts.join('');
};

const serializeLine3DChart = (chart: Line3DChart): string => {
  const parts: string[] = ['<c:line3DChart>', `<c:grouping val="${chart.grouping}"/>`];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeLineSeries(s));
  if (chart.gapDepth !== undefined) parts.push(`<c:gapDepth val="${chart.gapDepth}"/>`);
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[2]}"/>`);
  parts.push('</c:line3DChart>');
  return parts.join('');
};

const serializePie3DChart = (chart: Pie3DChart): string => {
  const parts: string[] = ['<c:pie3DChart>'];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push('</c:pie3DChart>');
  return parts.join('');
};

const serializeArea3DChart = (chart: Area3DChart): string => {
  const parts: string[] = ['<c:area3DChart>', `<c:grouping val="${chart.grouping}"/>`];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  if (chart.gapDepth !== undefined) parts.push(`<c:gapDepth val="${chart.gapDepth}"/>`);
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[2]}"/>`);
  parts.push('</c:area3DChart>');
  return parts.join('');
};

const serializeSurface3DChart = (chart: Surface3DChart): string => {
  const parts: string[] = ['<c:surface3DChart>'];
  if (chart.wireframe !== undefined) parts.push(`<c:wireframe val="${chart.wireframe ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[2]}"/>`);
  parts.push('</c:surface3DChart>');
  return parts.join('');
};

const serializeAxis = (tag: 'catAx' | 'valAx', ax: CategoryAxis | ValueAxis): string => {
  const parts: string[] = [
    `<c:${tag}>`,
    `<c:axId val="${ax.axId}"/>`,
    '<c:scaling><c:orientation val="minMax"/></c:scaling>',
    `<c:delete val="${ax.delete ? '1' : '0'}"/>`,
    `<c:axPos val="${ax.position ?? (tag === 'catAx' ? 'b' : 'l')}"/>`,
  ];
  if (tag === 'valAx' && (ax as ValueAxis).majorGridlines) parts.push('<c:majorGridlines/>');
  // ECMA-376 element order places spPr / txPr immediately before crossAx.
  if (ax.spPr) parts.push(serializeShapeProperties(ax.spPr));
  if (ax.txPr) parts.push(serializeTextBody(ax.txPr, 'c:txPr'));
  parts.push(`<c:crossAx val="${ax.crossAx}"/>`);
  parts.push(`</c:${tag}>`);
  return parts.join('');
};

const serializeChartTitle = (title: ChartTitle): string => {
  const parts: string[] = ['<c:title>'];
  if (title.tx) {
    // Wrap a TextBody as <c:tx><c:rich>...</c:rich></c:tx>.
    parts.push('<c:tx>');
    parts.push(serializeTextBody(title.tx, 'c:rich'));
    parts.push('</c:tx>');
  } else if (title.text !== undefined) {
    parts.push(
      '<c:tx>',
      '<c:rich>',
      '<a:bodyPr/><a:lstStyle/><a:p>',
      `<a:r><a:t>${escapeText(title.text)}</a:t></a:r>`,
      '</a:p>',
      '</c:rich>',
      '</c:tx>',
    );
  }
  if (title.overlay !== undefined) {
    parts.push(`<c:overlay val="${title.overlay ? '1' : '0'}"/>`);
  } else {
    parts.push('<c:overlay val="0"/>');
  }
  if (title.spPr) parts.push(serializeShapeProperties(title.spPr));
  if (title.txPr) parts.push(serializeTextBody(title.txPr, 'c:txPr'));
  parts.push('</c:title>');
  return parts.join('');
};

const serializeChartKind = (chart: ChartKind): string => {
  switch (chart.kind) {
    case 'bar':
      return serializeBarChart(chart);
    case 'line':
      return serializeLineChart(chart);
    case 'area':
      return serializeAreaChart(chart);
    case 'pie':
      return serializePieChart(chart);
    case 'doughnut':
      return serializeDoughnutChart(chart);
    case 'scatter':
      return serializeScatterChart(chart);
    case 'radar':
      return serializeRadarChart(chart);
    case 'bubble':
      return serializeBubbleChart(chart);
    case 'stock':
      return serializeStockChart(chart);
    case 'surface':
      return serializeSurfaceChart(chart);
    case 'ofPie':
      return serializeOfPieChart(chart);
    case 'bar3D':
      return serializeBar3DChart(chart);
    case 'line3D':
      return serializeLine3DChart(chart);
    case 'pie3D':
      return serializePie3DChart(chart);
    case 'area3D':
      return serializeArea3DChart(chart);
    case 'surface3D':
      return serializeSurface3DChart(chart);
  }
};

const serializePlotArea = (plotArea: PlotArea): string => {
  const parts: string[] = ['<c:plotArea>', '<c:layout/>'];
  parts.push(serializeChartKind(plotArea.chart));
  if (plotArea.catAx) parts.push(serializeAxis('catAx', plotArea.catAx));
  if (plotArea.valAx) parts.push(serializeAxis('valAx', plotArea.valAx));
  if (plotArea.spPr) parts.push(serializeShapeProperties(plotArea.spPr));
  parts.push('</c:plotArea>');
  return parts.join('');
};

const serializeLegend = (legend: Legend): string => {
  const parts: string[] = ['<c:legend>', `<c:legendPos val="${legend.position}"/>`];
  if (legend.overlay !== undefined) parts.push(`<c:overlay val="${legend.overlay ? '1' : '0'}"/>`);
  if (legend.spPr) parts.push(serializeShapeProperties(legend.spPr));
  if (legend.txPr) parts.push(serializeTextBody(legend.txPr, 'c:txPr'));
  parts.push('</c:legend>');
  return parts.join('');
};

/** Serialise a ChartSpace to its `xl/charts/chartN.xml` bytes. */
export function chartToBytes(space: ChartSpace): Uint8Array {
  return new TextEncoder().encode(serializeChartSpace(space));
}

export function serializeChartSpace(space: ChartSpace): string {
  const parts: string[] = [
    XML_HEADER,
    `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_NS}">`,
    '<c:chart>',
  ];
  if (space.title !== undefined) parts.push(serializeChartTitle(space.title));
  // openpyxl emits autoTitleDeleted="0" for charts that have a title; we
  // skip it for stage-1 since Excel tolerates the absence.
  parts.push(serializePlotArea(space.plotArea));
  if (space.legend) parts.push(serializeLegend(space.legend));
  if (space.plotVisOnly !== undefined) parts.push(`<c:plotVisOnly val="${space.plotVisOnly ? '1' : '0'}"/>`);
  if (space.dispBlanksAs !== undefined) parts.push(`<c:dispBlanksAs val="${space.dispBlanksAs}"/>`);
  parts.push('</c:chart>');
  // chartSpace-level spPr / txPr are siblings of <c:chart>, emitted after.
  if (space.spPr) parts.push(serializeShapeProperties(space.spPr));
  if (space.txPr) parts.push(serializeTextBody(space.txPr, 'c:txPr'));
  parts.push('</c:chartSpace>');
  return parts.join('');
}

// SHEET_DRAWING_NS is imported only to keep import surface stable; the
// chart serialiser doesn't need it directly.
void SHEET_DRAWING_NS;
