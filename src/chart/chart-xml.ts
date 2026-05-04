// xl/charts/chartN.xml read/write. Per docs/plan/08-charts-drawings.md §5.
//
// Stage-1 covers BarChart end-to-end: parse + serialize with title /
// legend / catAx / valAx / series (cat + val refs + numCache /
// strCache). Other chart kinds slot in alongside as their own
// `<c:lineChart>` / `<c:pieChart>` / etc. parsers.

import { OpenXmlSchemaError } from '../utils/exceptions';
import { CHART_NS, REL_NS, SHEET_DRAWING_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, findChildren, type XmlNode } from '../xml/tree';
import {
  type BarChart,
  type BarDirection,
  type BarSeries,
  type CategoryAxis,
  type CategoryRef,
  type ChartSpace,
  type GroupingType,
  type Legend,
  type LegendPosition,
  makeBarChart,
  makeBarSeries,
  makeChartSpace,
  type NumericRef,
  type PlotArea,
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
const AX_POS_TAG = `{${CHART_NS}}axPos`;
const CROSS_AX_TAG = `{${CHART_NS}}crossAx`;
const MAJOR_GRIDLINES_TAG = `{${CHART_NS}}majorGridlines`;
const LEGEND_TAG = `{${CHART_NS}}legend`;
const LEGEND_POS_TAG = `{${CHART_NS}}legendPos`;
const PLOT_VIS_ONLY_TAG = `{${CHART_NS}}plotVisOnly`;
const DISP_BLANKS_AS_TAG = `{${CHART_NS}}dispBlanksAs`;
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

const parseTitleString = (titleEl: XmlNode): string | undefined => {
  // Excel surfaces the title text inside <c:tx><c:rich>...<a:r><a:t>text</a:t></a:r></c:rich></c:tx>.
  const tx = findChild(titleEl, TX_TAG);
  if (!tx) return undefined;
  const rich = findChild(tx, RICH_TAG);
  if (!rich) return undefined;
  let out = '';
  for (const p of rich.children) {
    for (const r of findChildren(p, A_R_TAG)) {
      const t = findChild(r, A_T_TAG);
      if (t?.text) out += t.text;
    }
  }
  return out.length > 0 ? out : undefined;
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
  return makeBarSeries(opts);
};

const parseBarChart = (barEl: XmlNode): BarChart => {
  const barDir = (valAttr(findChild(barEl, BAR_DIR_TAG)) ?? 'col') as BarDirection;
  const grouping = (valAttr(findChild(barEl, GROUPING_TAG)) ?? 'clustered') as GroupingType;
  const varyColors = boolVal(findChild(barEl, VARY_COLORS_TAG));
  const gapWidth = intVal(findChild(barEl, GAP_WIDTH_TAG));
  const series: BarSeries[] = [];
  for (const ser of findChildren(barEl, SER_TAG)) {
    const s = parseSeries(ser);
    if (s) series.push(s);
  }
  const axIdNodes = findChildren(barEl, AX_ID_TAG);
  const axIds: [number, number] = [intVal(axIdNodes[0]) ?? 1, intVal(axIdNodes[1]) ?? 2];
  return makeBarChart({
    barDir,
    grouping,
    series,
    axIds,
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(gapWidth !== undefined ? { gapWidth } : {}),
  });
};

const parseAxis = (
  axEl: XmlNode,
): { axId: number; crossAx: number; position?: 'b' | 't' | 'l' | 'r'; delete?: boolean; majorGridlines?: boolean } => {
  const axId = intVal(findChild(axEl, AX_ID_TAG)) ?? 0;
  const crossAx = intVal(findChild(axEl, CROSS_AX_TAG)) ?? 0;
  const positionRaw = valAttr(findChild(axEl, AX_POS_TAG));
  const validPos = positionRaw === 'b' || positionRaw === 't' || positionRaw === 'l' || positionRaw === 'r';
  const del = boolVal(findChild(axEl, DELETE_TAG));
  const majorGridlines = findChild(axEl, MAJOR_GRIDLINES_TAG) !== undefined ? true : undefined;
  return {
    axId,
    crossAx,
    ...(validPos ? { position: positionRaw as 'b' | 't' | 'l' | 'r' } : {}),
    ...(del !== undefined ? { delete: del } : {}),
    ...(majorGridlines !== undefined ? { majorGridlines } : {}),
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
  const barEl = findChild(plotAreaEl, BAR_CHART_TAG);
  if (!barEl) throw new OpenXmlSchemaError('parseChartXml: stage-1 only supports <barChart>');
  const chart = parseBarChart(barEl);
  const catAxEl = findChild(plotAreaEl, CAT_AX_TAG);
  const valAxEl = findChild(plotAreaEl, VAL_AX_TAG);
  const plotArea: PlotArea = {
    chart,
    ...(catAxEl ? { catAx: parseAxis(catAxEl) as CategoryAxis } : {}),
    ...(valAxEl ? { valAx: parseAxis(valAxEl) as ValueAxis } : {}),
  };
  const titleEl = findChild(chartEl, TITLE_TAG);
  const title = titleEl ? parseTitleString(titleEl) : undefined;
  const legendEl = findChild(chartEl, LEGEND_TAG);
  let legend: Legend | undefined;
  if (legendEl) {
    const posRaw = valAttr(findChild(legendEl, LEGEND_POS_TAG)) as LegendPosition | undefined;
    legend = { position: posRaw ?? 'r' };
  }
  const plotVisOnly = boolVal(findChild(chartEl, PLOT_VIS_ONLY_TAG));
  const dispBlanksAs = valAttr(findChild(chartEl, DISP_BLANKS_AS_TAG)) as ChartSpace['dispBlanksAs'];
  return makeChartSpace({
    plotArea,
    ...(title !== undefined ? { title } : {}),
    ...(legend ? { legend } : {}),
    ...(plotVisOnly !== undefined ? { plotVisOnly } : {}),
    ...(dispBlanksAs ? { dispBlanksAs } : {}),
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

const serializeAxis = (tag: 'catAx' | 'valAx', ax: CategoryAxis | ValueAxis): string => {
  const parts: string[] = [
    `<c:${tag}>`,
    `<c:axId val="${ax.axId}"/>`,
    '<c:scaling><c:orientation val="minMax"/></c:scaling>',
    `<c:delete val="${ax.delete ? '1' : '0'}"/>`,
    `<c:axPos val="${ax.position ?? (tag === 'catAx' ? 'b' : 'l')}"/>`,
  ];
  if (tag === 'valAx' && (ax as ValueAxis).majorGridlines) parts.push('<c:majorGridlines/>');
  parts.push(`<c:crossAx val="${ax.crossAx}"/>`);
  parts.push(`</c:${tag}>`);
  return parts.join('');
};

const serializeTitle = (title: string): string =>
  [
    '<c:title>',
    '<c:tx>',
    '<c:rich>',
    '<a:bodyPr/><a:lstStyle/><a:p>',
    `<a:r><a:t>${escapeText(title)}</a:t></a:r>`,
    '</a:p>',
    '</c:rich>',
    '</c:tx>',
    '<c:overlay val="0"/>',
    '</c:title>',
  ].join('');

const serializePlotArea = (plotArea: PlotArea): string => {
  const parts: string[] = ['<c:plotArea>', '<c:layout/>'];
  if (plotArea.chart.kind === 'bar') parts.push(serializeBarChart(plotArea.chart));
  if (plotArea.catAx) parts.push(serializeAxis('catAx', plotArea.catAx));
  if (plotArea.valAx) parts.push(serializeAxis('valAx', plotArea.valAx));
  parts.push('</c:plotArea>');
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
  if (space.title !== undefined) parts.push(serializeTitle(space.title));
  // openpyxl emits autoTitleDeleted="0" for charts that have a title; we
  // skip it for stage-1 since Excel tolerates the absence.
  parts.push(serializePlotArea(space.plotArea));
  if (space.legend) {
    parts.push(`<c:legend><c:legendPos val="${space.legend.position}"/></c:legend>`);
  }
  if (space.plotVisOnly !== undefined) parts.push(`<c:plotVisOnly val="${space.plotVisOnly ? '1' : '0'}"/>`);
  if (space.dispBlanksAs !== undefined) parts.push(`<c:dispBlanksAs val="${space.dispBlanksAs}"/>`);
  parts.push('</c:chart></c:chartSpace>');
  return parts.join('');
}

// SHEET_DRAWING_NS is imported only to keep import surface stable; the
// chart serialiser doesn't need it directly.
void SHEET_DRAWING_NS;
