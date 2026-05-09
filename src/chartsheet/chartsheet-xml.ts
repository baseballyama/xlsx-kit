// `xl/chartsheets/sheetN.xml` reader / writer.

import { OpenXmlSchemaError } from '../utils/exceptions';
import { REL_NS, SHEET_MAIN_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, findChildren, type XmlNode } from '../xml/tree';
import {
  parseHeaderFooter,
  parsePageMargins,
  parsePageSetup,
  parseWebPublishItem,
} from '../worksheet/reader';
import {
  serializeHeaderFooter,
  serializePageMargins,
  serializePageSetup,
  serializeWebPublishItems,
} from '../worksheet/writer';
import {
  type Chartsheet,
  type ChartsheetCustomSheetView,
  type ChartsheetDrawingHF,
  type ChartsheetProperties,
  type ChartsheetProtection,
  type ChartsheetView,
  makeChartsheet,
} from './chartsheet';

const CHARTSHEET_TAG = `{${SHEET_MAIN_NS}}chartsheet`;
const SHEET_PR_TAG = `{${SHEET_MAIN_NS}}sheetPr`;
const SHEET_VIEWS_TAG = `{${SHEET_MAIN_NS}}sheetViews`;
const SHEET_VIEW_TAG = `{${SHEET_MAIN_NS}}sheetView`;
const SHEET_PROTECTION_TAG = `{${SHEET_MAIN_NS}}sheetProtection`;
const TAB_COLOR_TAG = `{${SHEET_MAIN_NS}}tabColor`;
const DRAWING_TAG = `{${SHEET_MAIN_NS}}drawing`;
const PAGE_MARGINS_TAG = `{${SHEET_MAIN_NS}}pageMargins`;
const PAGE_SETUP_TAG = `{${SHEET_MAIN_NS}}pageSetup`;
const HEADER_FOOTER_TAG = `{${SHEET_MAIN_NS}}headerFooter`;
const LEGACY_DRAWING_TAG = `{${SHEET_MAIN_NS}}legacyDrawing`;
const LEGACY_DRAWING_HF_TAG = `{${SHEET_MAIN_NS}}legacyDrawingHF`;
const DRAWING_HF_TAG = `{${SHEET_MAIN_NS}}drawingHF`;
const PICTURE_TAG = `{${SHEET_MAIN_NS}}picture`;
const WEB_PUBLISH_ITEMS_TAG = `{${SHEET_MAIN_NS}}webPublishItems`;
const WEB_PUBLISH_ITEM_TAG = `{${SHEET_MAIN_NS}}webPublishItem`;
const CUSTOM_SHEET_VIEWS_TAG = `{${SHEET_MAIN_NS}}customSheetViews`;
const CUSTOM_SHEET_VIEW_TAG = `{${SHEET_MAIN_NS}}customSheetView`;

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const escapeAttr = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const parseBool = (v: string | undefined): boolean | undefined => {
  if (v === undefined) return undefined;
  if (v === '1' || v === 'true') return true;
  if (v === '0' || v === 'false') return false;
  return undefined;
};

const parseInt10 = (v: string | undefined): number | undefined => {
  if (v === undefined) return undefined;
  const n = Number.parseInt(v, 10);
  return Number.isInteger(n) ? n : undefined;
};

const parseSheetPr = (el: XmlNode): ChartsheetProperties | undefined => {
  const out: ChartsheetProperties = {};
  const published = parseBool(el.attrs['published']);
  if (published !== undefined) out.published = published;
  if (el.attrs['codeName'] !== undefined) out.codeName = el.attrs['codeName'];
  const tab = findChild(el, TAB_COLOR_TAG);
  if (tab) {
    const rgb = tab.attrs['rgb'];
    if (rgb) out.tabColorRgb = rgb.toUpperCase();
  }
  return Object.keys(out).length > 0 ? out : undefined;
};

const parseSheetView = (el: XmlNode): ChartsheetView => {
  const workbookViewId = parseInt10(el.attrs['workbookViewId']) ?? 0;
  const out: ChartsheetView = { workbookViewId };
  const tabSelected = parseBool(el.attrs['tabSelected']);
  if (tabSelected !== undefined) out.tabSelected = tabSelected;
  const zoomScale = parseInt10(el.attrs['zoomScale']);
  if (zoomScale !== undefined) out.zoomScale = zoomScale;
  const zoomToFit = parseBool(el.attrs['zoomToFit']);
  if (zoomToFit !== undefined) out.zoomToFit = zoomToFit;
  return out;
};

const parseSheetProtection = (el: XmlNode): ChartsheetProtection => {
  const out: ChartsheetProtection = {};
  const content = parseBool(el.attrs['content']);
  if (content !== undefined) out.content = content;
  const objects = parseBool(el.attrs['objects']);
  if (objects !== undefined) out.objects = objects;
  if (el.attrs['algorithmName'] !== undefined) out.algorithmName = el.attrs['algorithmName'];
  if (el.attrs['hashValue'] !== undefined) out.hashValue = el.attrs['hashValue'];
  if (el.attrs['saltValue'] !== undefined) out.saltValue = el.attrs['saltValue'];
  const spinCount = parseInt10(el.attrs['spinCount']);
  if (spinCount !== undefined) out.spinCount = spinCount;
  return out;
};

/**
 * Parse a chartsheet part. Returns a Chartsheet with `title` set to the
 * provided fallback (Excel doesn't store the display name inside the chartsheet
 * part itself — it lives in workbook.xml's `<sheet name>`).
 */
export function parseChartsheetXml(bytes: Uint8Array | string, title: string): Chartsheet {
  const root = parseXml(bytes);
  if (root.name !== CHARTSHEET_TAG) {
    throw new OpenXmlSchemaError(`parseChartsheetXml: root is "${root.name}", expected ${CHARTSHEET_TAG}`);
  }
  const cs = makeChartsheet(title);
  const sheetPr = findChild(root, SHEET_PR_TAG);
  if (sheetPr) {
    const props = parseSheetPr(sheetPr);
    if (props) cs.properties = props;
  }
  const sheetViews = findChild(root, SHEET_VIEWS_TAG);
  if (sheetViews) {
    const views: ChartsheetView[] = [];
    for (const v of findChildren(sheetViews, SHEET_VIEW_TAG)) views.push(parseSheetView(v));
    if (views.length > 0) cs.views = views;
  }
  const protectionEl = findChild(root, SHEET_PROTECTION_TAG);
  if (protectionEl) cs.protection = parseSheetProtection(protectionEl);

  const csvEl = findChild(root, CUSTOM_SHEET_VIEWS_TAG);
  if (csvEl) {
    for (const v of findChildren(csvEl, CUSTOM_SHEET_VIEW_TAG)) {
      const view = parseChartsheetCustomSheetView(v);
      if (view) cs.customSheetViews.push(view);
    }
  }

  const pmEl = findChild(root, PAGE_MARGINS_TAG);
  if (pmEl) {
    const pm = parsePageMargins(pmEl);
    if (pm) cs.pageMargins = pm;
  }
  const psEl = findChild(root, PAGE_SETUP_TAG);
  if (psEl) {
    const ps = parsePageSetup(psEl);
    if (ps) cs.pageSetup = ps;
  }
  const hfEl = findChild(root, HEADER_FOOTER_TAG);
  if (hfEl) {
    const hf = parseHeaderFooter(hfEl);
    if (hf) cs.headerFooter = hf;
  }

  const ldEl = findChild(root, LEGACY_DRAWING_TAG);
  if (ldEl) {
    const rId = ldEl.attrs[`{${REL_NS}}id`];
    if (rId) cs.legacyDrawingRId = rId;
  }
  const ldhfEl = findChild(root, LEGACY_DRAWING_HF_TAG);
  if (ldhfEl) {
    const rId = ldhfEl.attrs[`{${REL_NS}}id`];
    if (rId) cs.legacyDrawingHFRId = rId;
  }
  const dhfEl = findChild(root, DRAWING_HF_TAG);
  if (dhfEl) {
    const dhf = parseDrawingHF(dhfEl);
    if (dhf) cs.drawingHF = dhf;
  }
  const picEl = findChild(root, PICTURE_TAG);
  if (picEl) {
    const rId = picEl.attrs[`{${REL_NS}}id`];
    if (rId) cs.backgroundPictureRId = rId;
  }
  const wpEl = findChild(root, WEB_PUBLISH_ITEMS_TAG);
  if (wpEl) {
    for (const wp of findChildren(wpEl, WEB_PUBLISH_ITEM_TAG)) {
      const item = parseWebPublishItem(wp);
      if (item) cs.webPublishItems.push(item);
    }
  }

  // Drawing reference — the actual rId / drawing payload is resolved by the
  // loader.
  return cs;
}

const CSV_STATE_VALUES: ReadonlyArray<NonNullable<ChartsheetCustomSheetView['state']>> = [
  'visible',
  'hidden',
  'veryHidden',
];

const parseChartsheetCustomSheetView = (
  el: XmlNode,
): ChartsheetCustomSheetView | undefined => {
  const guid = el.attrs['guid'];
  if (!guid) return undefined;
  const out: ChartsheetCustomSheetView = { guid };
  const scale = parseInt10(el.attrs['scale']);
  if (scale !== undefined) out.scale = scale;
  const stateRaw = el.attrs['state'];
  if (
    stateRaw &&
    CSV_STATE_VALUES.includes(stateRaw as NonNullable<ChartsheetCustomSheetView['state']>)
  ) {
    out.state = stateRaw as NonNullable<ChartsheetCustomSheetView['state']>;
  }
  const ztf = parseBool(el.attrs['zoomToFit']);
  if (ztf !== undefined) out.zoomToFit = ztf;
  const pmEl = findChild(el, PAGE_MARGINS_TAG);
  if (pmEl) {
    const pm = parsePageMargins(pmEl);
    if (pm) out.pageMargins = pm;
  }
  const psEl = findChild(el, PAGE_SETUP_TAG);
  if (psEl) {
    const ps = parsePageSetup(psEl);
    if (ps) out.pageSetup = ps;
  }
  const hfEl = findChild(el, HEADER_FOOTER_TAG);
  if (hfEl) {
    const hf = parseHeaderFooter(hfEl);
    if (hf) out.headerFooter = hf;
  }
  return out;
};

const DRAWING_HF_INT_KEYS = [
  'lho',
  'cho',
  'rho',
  'lhe',
  'che',
  'rhe',
  'lhf',
  'chf',
  'rhf',
  'lfo',
  'cfo',
  'rfo',
  'lfe',
  'cfe',
  'rfe',
  'lff',
  'cff',
  'rff',
] as const;

const parseDrawingHF = (el: XmlNode): ChartsheetDrawingHF | undefined => {
  const rId = el.attrs[`{${REL_NS}}id`];
  if (!rId) return undefined;
  const out: ChartsheetDrawingHF = { rId };
  for (const k of DRAWING_HF_INT_KEYS) {
    const raw = el.attrs[k];
    if (raw === undefined) continue;
    const n = Number.parseInt(raw, 10);
    if (Number.isInteger(n)) (out as unknown as Record<string, unknown>)[k] = n;
  }
  return out;
};

/**
 * Optional drawing rId injected by the writer. The chartsheet part itself only
 * references the drawing by relationship id; the writer supplies the id once it
 * has registered the drawing.
 */
export interface ChartsheetSerializeOptions {
  drawingRId?: string;
}

const serializeSheetPr = (p: ChartsheetProperties): string => {
  const attrs: string[] = [];
  if (p.published !== undefined) attrs.push(`published="${p.published ? '1' : '0'}"`);
  if (p.codeName !== undefined) attrs.push(`codeName="${escapeAttr(p.codeName)}"`);
  const tab = p.tabColorRgb ? `<tabColor rgb="${p.tabColorRgb}"/>` : '';
  return tab === ''
    ? `<sheetPr${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}/>`
    : `<sheetPr${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}>${tab}</sheetPr>`;
};

const serializeSheetView = (v: ChartsheetView): string => {
  const a: string[] = [`workbookViewId="${v.workbookViewId}"`];
  if (v.tabSelected !== undefined) a.push(`tabSelected="${v.tabSelected ? '1' : '0'}"`);
  if (v.zoomScale !== undefined) a.push(`zoomScale="${v.zoomScale}"`);
  if (v.zoomToFit !== undefined) a.push(`zoomToFit="${v.zoomToFit ? '1' : '0'}"`);
  return `<sheetView ${a.join(' ')}/>`;
};

const serializeSheetProtection = (p: ChartsheetProtection): string => {
  const a: string[] = [];
  if (p.content !== undefined) a.push(`content="${p.content ? '1' : '0'}"`);
  if (p.objects !== undefined) a.push(`objects="${p.objects ? '1' : '0'}"`);
  if (p.algorithmName !== undefined) a.push(`algorithmName="${escapeAttr(p.algorithmName)}"`);
  if (p.hashValue !== undefined) a.push(`hashValue="${escapeAttr(p.hashValue)}"`);
  if (p.saltValue !== undefined) a.push(`saltValue="${escapeAttr(p.saltValue)}"`);
  if (p.spinCount !== undefined) a.push(`spinCount="${p.spinCount}"`);
  return `<sheetProtection${a.length > 0 ? ` ${a.join(' ')}` : ''}/>`;
};

export function serializeChartsheet(cs: Chartsheet, opts: ChartsheetSerializeOptions = {}): string {
  const parts: string[] = [XML_HEADER, `<chartsheet xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}">`];
  if (cs.properties) parts.push(serializeSheetPr(cs.properties));
  parts.push('<sheetViews>');
  for (const v of cs.views) parts.push(serializeSheetView(v));
  parts.push('</sheetViews>');
  if (cs.protection) parts.push(serializeSheetProtection(cs.protection));
  if (cs.customSheetViews.length > 0) parts.push(serializeChartsheetCustomSheetViews(cs.customSheetViews));
  if (cs.pageMargins) parts.push(serializePageMargins(cs.pageMargins));
  if (cs.pageSetup) {
    const ps = serializePageSetup(cs.pageSetup);
    if (ps) parts.push(ps);
  }
  if (cs.headerFooter) {
    const hf = serializeHeaderFooter(cs.headerFooter);
    if (hf) parts.push(hf);
  }
  if (opts.drawingRId !== undefined) {
    parts.push(`<drawing r:id="${escapeAttr(opts.drawingRId)}"/>`);
  }
  if (cs.legacyDrawingRId !== undefined) {
    parts.push(`<legacyDrawing r:id="${escapeAttr(cs.legacyDrawingRId)}"/>`);
  }
  if (cs.legacyDrawingHFRId !== undefined) {
    parts.push(`<legacyDrawingHF r:id="${escapeAttr(cs.legacyDrawingHFRId)}"/>`);
  }
  if (cs.drawingHF) {
    parts.push(serializeDrawingHF(cs.drawingHF));
  }
  if (cs.backgroundPictureRId !== undefined) {
    parts.push(`<picture r:id="${escapeAttr(cs.backgroundPictureRId)}"/>`);
  }
  if (cs.webPublishItems.length > 0) parts.push(serializeWebPublishItems(cs.webPublishItems));
  parts.push('</chartsheet>');
  return parts.join('');
}

const serializeChartsheetCustomSheetViews = (
  views: ReadonlyArray<ChartsheetCustomSheetView>,
): string => {
  const parts: string[] = ['<customSheetViews>'];
  for (const v of views) {
    let attrs = ` guid="${escapeAttr(v.guid)}"`;
    if (v.scale !== undefined) attrs += ` scale="${v.scale}"`;
    if (v.state !== undefined) attrs += ` state="${v.state}"`;
    if (v.zoomToFit !== undefined) attrs += ` zoomToFit="${v.zoomToFit ? '1' : '0'}"`;

    const inner: string[] = [];
    if (v.pageMargins) inner.push(serializePageMargins(v.pageMargins));
    if (v.pageSetup) {
      const ps = serializePageSetup(v.pageSetup);
      if (ps) inner.push(ps);
    }
    if (v.headerFooter) {
      const hf = serializeHeaderFooter(v.headerFooter);
      if (hf) inner.push(hf);
    }
    if (inner.length === 0) parts.push(`<customSheetView${attrs}/>`);
    else parts.push(`<customSheetView${attrs}>${inner.join('')}</customSheetView>`);
  }
  parts.push('</customSheetViews>');
  return parts.join('');
};

const serializeDrawingHF = (dhf: ChartsheetDrawingHF): string => {
  let attrs = ` r:id="${escapeAttr(dhf.rId)}"`;
  for (const k of DRAWING_HF_INT_KEYS) {
    const v = (dhf as unknown as Record<string, number | undefined>)[k];
    if (v !== undefined) attrs += ` ${k}="${v}"`;
  }
  return `<drawingHF${attrs}/>`;
};

export function chartsheetToBytes(cs: Chartsheet, opts: ChartsheetSerializeOptions = {}): Uint8Array {
  return new TextEncoder().encode(serializeChartsheet(cs, opts));
}

void DRAWING_TAG;
