// `xl/drawings/chartDrawingN.xml` reader/writer. Per docs/plan/08-charts-drawings.md §8.

import {
  parseShapeProperties,
  parseTextBody,
  serializeShapeProperties,
  serializeTextBody,
} from '../drawing/dml/dml-xml';
import type { PositiveSize2D } from '../drawing/dml/shape-properties';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { CHART_DRAWING_NS, DRAWING_NS, REL_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, type XmlNode } from '../xml/tree';
import {
  type ChartDrawing,
  type ChartDrawingPicture,
  type ChartDrawingShape,
  type ChartRelativeMarker,
  makeChartDrawing,
  type UserShapeAnchor,
  type UserShapeContent,
} from './user-shapes';

const C = (local: string): string => `{${CHART_DRAWING_NS}}${local}`;
const A = (local: string): string => `{${DRAWING_NS}}${local}`;

const USER_SHAPES_TAG = C('userShapes');
const REL_SIZE_ANCHOR = C('relSizeAnchor');
const ABS_SIZE_ANCHOR = C('absSizeAnchor');
const FROM = C('from');
const TO = C('to');
const EXT = C('ext');
const X_TAG = C('x');
const Y_TAG = C('y');
const SP = C('sp');
const NV_SP_PR = C('nvSpPr');
const C_NV_PR = C('cNvPr');
const C_NV_SP_PR = C('cNvSpPr');
const SP_PR = C('spPr');
const TX_BODY = C('txBody');
const PIC = C('pic');
const NV_PIC_PR = C('nvPicPr');
const BLIP_FILL = C('blipFill');
const A_BLIP = A('blip');

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const escapeText = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttr = (s: string): string => escapeText(s).replace(/"/g, '&quot;');

const parseFloatText = (text: string | undefined): number | undefined => {
  if (text === undefined) return undefined;
  const n = Number.parseFloat(text);
  return Number.isFinite(n) ? n : undefined;
};

const parseMarker = (el: XmlNode): ChartRelativeMarker => ({
  x: parseFloatText(findChild(el, X_TAG)?.text) ?? 0,
  y: parseFloatText(findChild(el, Y_TAG)?.text) ?? 0,
});

const parseExt = (el: XmlNode): PositiveSize2D => {
  const cx = Number.parseInt(el.attrs['cx'] ?? '0', 10);
  const cy = Number.parseInt(el.attrs['cy'] ?? '0', 10);
  return { cx: Number.isFinite(cx) ? cx : 0, cy: Number.isFinite(cy) ? cy : 0 };
};

const parseCNvPr = (
  cNvPr: XmlNode | undefined,
): { id: number; name?: string; descr?: string; hidden?: boolean } => {
  if (!cNvPr) return { id: 0 };
  const id = Number.parseInt(cNvPr.attrs['id'] ?? '0', 10);
  const out: { id: number; name?: string; descr?: string; hidden?: boolean } = {
    id: Number.isFinite(id) ? id : 0,
  };
  if (cNvPr.attrs['name'] !== undefined) out.name = cNvPr.attrs['name'];
  if (cNvPr.attrs['descr'] !== undefined) out.descr = cNvPr.attrs['descr'];
  const hiddenRaw = cNvPr.attrs['hidden'];
  if (hiddenRaw === '1' || hiddenRaw === 'true') out.hidden = true;
  return out;
};

const parseShape = (el: XmlNode): ChartDrawingShape => {
  const nvSpPr = findChild(el, NV_SP_PR);
  const cNvPr = nvSpPr ? findChild(nvSpPr, C_NV_PR) : undefined;
  const cNvSpPr = nvSpPr ? findChild(nvSpPr, C_NV_SP_PR) : undefined;
  const txBoxRaw = cNvSpPr?.attrs['txBox'];
  const out: ChartDrawingShape = parseCNvPr(cNvPr);
  if (txBoxRaw === '1' || txBoxRaw === 'true') out.txBox = true;
  const spPrEl = findChild(el, SP_PR);
  if (spPrEl) out.spPr = parseShapeProperties(spPrEl);
  const txBodyEl = findChild(el, TX_BODY);
  if (txBodyEl) out.txBody = parseTextBody(txBodyEl);
  return out;
};

const parsePicture = (el: XmlNode): ChartDrawingPicture => {
  const nvPicPr = findChild(el, NV_PIC_PR);
  const cNvPr = nvPicPr ? findChild(nvPicPr, C_NV_PR) : undefined;
  const out: ChartDrawingPicture = parseCNvPr(cNvPr);
  const blipFill = findChild(el, BLIP_FILL);
  if (blipFill) {
    const blip = findChild(blipFill, A_BLIP);
    const embed = blip?.attrs[`{${REL_NS}}embed`];
    if (embed) out.embedRId = embed;
  }
  const spPrEl = findChild(el, SP_PR);
  if (spPrEl) out.spPr = parseShapeProperties(spPrEl);
  return out;
};

const parseContent = (el: XmlNode): UserShapeContent | undefined => {
  const sp = findChild(el, SP);
  if (sp) return { kind: 'shape', shape: parseShape(sp) };
  const pic = findChild(el, PIC);
  if (pic) return { kind: 'picture', picture: parsePicture(pic) };
  return undefined;
};

const parseAnchor = (el: XmlNode): UserShapeAnchor | undefined => {
  const fromEl = findChild(el, FROM);
  if (!fromEl) return undefined;
  const from = parseMarker(fromEl);
  const content = parseContent(el);
  if (!content) return undefined;
  if (el.name === REL_SIZE_ANCHOR) {
    const toEl = findChild(el, TO);
    if (!toEl) return undefined;
    return { kind: 'relSize', from, to: parseMarker(toEl), content };
  }
  if (el.name === ABS_SIZE_ANCHOR) {
    const extEl = findChild(el, EXT);
    if (!extEl) return undefined;
    return { kind: 'absSize', from, ext: parseExt(extEl), content };
  }
  return undefined;
};

export function parseUserShapesXml(bytes: Uint8Array | string): ChartDrawing {
  const root = parseXml(bytes);
  if (root.name !== USER_SHAPES_TAG) {
    throw new OpenXmlSchemaError(`parseUserShapesXml: root is "${root.name}", expected ${USER_SHAPES_TAG}`);
  }
  const shapes: UserShapeAnchor[] = [];
  for (const child of root.children) {
    if (typeof child === 'string') continue;
    if (child.name !== REL_SIZE_ANCHOR && child.name !== ABS_SIZE_ANCHOR) continue;
    const anchor = parseAnchor(child);
    if (anchor) shapes.push(anchor);
  }
  return makeChartDrawing(shapes);
}

const serializeMarker = (tag: string, m: ChartRelativeMarker): string =>
  `<cdr:${tag}><cdr:x>${m.x}</cdr:x><cdr:y>${m.y}</cdr:y></cdr:${tag}>`;

const serializeShape = (s: ChartDrawingShape): string => {
  const cNvPrAttrs: string[] = [`id="${s.id}"`, `name="${escapeAttr(s.name ?? `Shape ${s.id}`)}"`];
  if (s.descr !== undefined) cNvPrAttrs.push(`descr="${escapeAttr(s.descr)}"`);
  if (s.hidden) cNvPrAttrs.push('hidden="1"');
  const cNvSpPr = s.txBox ? '<cdr:cNvSpPr txBox="1"/>' : '<cdr:cNvSpPr/>';
  const spPr = s.spPr
    ? serializeShapeProperties(s.spPr, 'cdr:spPr')
    : '<cdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></cdr:spPr>';
  const txBody = s.txBody ? serializeTextBody(s.txBody, 'cdr:txBody') : '';
  return [
    '<cdr:sp macro="" textlink="">',
    `<cdr:nvSpPr><cdr:cNvPr ${cNvPrAttrs.join(' ')}/>${cNvSpPr}</cdr:nvSpPr>`,
    spPr,
    txBody,
    '</cdr:sp>',
  ].join('');
};

const serializePicture = (p: ChartDrawingPicture): string => {
  const cNvPrAttrs: string[] = [`id="${p.id}"`, `name="${escapeAttr(p.name ?? `Picture ${p.id}`)}"`];
  if (p.descr !== undefined) cNvPrAttrs.push(`descr="${escapeAttr(p.descr)}"`);
  const blip = p.embedRId
    ? `<a:blip xmlns:r="${REL_NS}" r:embed="${escapeAttr(p.embedRId)}"/>`
    : '<a:blip/>';
  const spPr = p.spPr
    ? serializeShapeProperties(p.spPr, 'cdr:spPr')
    : '<cdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></cdr:spPr>';
  return [
    '<cdr:pic>',
    `<cdr:nvPicPr><cdr:cNvPr ${cNvPrAttrs.join(' ')}/><cdr:cNvPicPr/></cdr:nvPicPr>`,
    `<cdr:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></cdr:blipFill>`,
    spPr,
    '</cdr:pic>',
  ].join('');
};

const serializeContent = (c: UserShapeContent): string =>
  c.kind === 'shape' ? serializeShape(c.shape) : serializePicture(c.picture);

const serializeAnchor = (a: UserShapeAnchor): string => {
  if (a.kind === 'relSize') {
    return [
      '<cdr:relSizeAnchor>',
      serializeMarker('from', a.from),
      serializeMarker('to', a.to),
      serializeContent(a.content),
      '</cdr:relSizeAnchor>',
    ].join('');
  }
  return [
    '<cdr:absSizeAnchor>',
    serializeMarker('from', a.from),
    `<cdr:ext cx="${a.ext.cx}" cy="${a.ext.cy}"/>`,
    serializeContent(a.content),
    '</cdr:absSizeAnchor>',
  ].join('');
};

export function serializeUserShapes(d: ChartDrawing): string {
  const parts: string[] = [
    XML_HEADER,
    `<cdr:userShapes xmlns:cdr="${CHART_DRAWING_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${REL_NS}">`,
  ];
  for (const a of d.shapes) parts.push(serializeAnchor(a));
  parts.push('</cdr:userShapes>');
  return parts.join('');
}

export function userShapesToBytes(d: ChartDrawing): Uint8Array {
  return new TextEncoder().encode(serializeUserShapes(d));
}
