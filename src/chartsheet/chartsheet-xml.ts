// `xl/chartsheets/sheetN.xml` reader / writer. Per docs/plan/08-charts-drawings.md §7.

import { OpenXmlSchemaError } from '../utils/exceptions';
import { REL_NS, SHEET_MAIN_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, findChildren, type XmlNode } from '../xml/tree';
import {
  type Chartsheet,
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
 * provided fallback (Excel doesn't store the display name inside the
 * chartsheet part itself — it lives in workbook.xml's `<sheet name>`).
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
  // Drawing reference — the actual rId / drawing payload is resolved by the loader.
  return cs;
}

/**
 * Optional drawing rId injected by the writer. The chartsheet part
 * itself only references the drawing by relationship id; the writer
 * supplies the id once it has registered the drawing.
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
  if (opts.drawingRId !== undefined) {
    parts.push(`<drawing r:id="${escapeAttr(opts.drawingRId)}"/>`);
  }
  parts.push('</chartsheet>');
  return parts.join('');
}

export function chartsheetToBytes(cs: Chartsheet, opts: ChartsheetSerializeOptions = {}): Uint8Array {
  return new TextEncoder().encode(serializeChartsheet(cs, opts));
}

void DRAWING_TAG;
