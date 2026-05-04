// `xl/drawings/drawingN.xml` reader/writer. Stage-1 supports anchor
// envelope round-trip for the chart variant; picture / shape / connector /
// group remain unsupported placeholders for later iterations.

import { OpenXmlSchemaError } from '../utils/exceptions';
import { REL_NS, SHEET_DRAWING_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, type XmlNode } from '../xml/tree';
import type { AnchorMarker, DrawingAnchor, Point2D, PositiveSize2D } from './anchor';
import { type ChartReference, type Drawing, type DrawingItem, makeDrawing } from './drawing';

const WS_DRAWING_TAG = `{${SHEET_DRAWING_NS}}wsDr`;
const ABSOLUTE_ANCHOR_TAG = `{${SHEET_DRAWING_NS}}absoluteAnchor`;
const ONE_CELL_ANCHOR_TAG = `{${SHEET_DRAWING_NS}}oneCellAnchor`;
const TWO_CELL_ANCHOR_TAG = `{${SHEET_DRAWING_NS}}twoCellAnchor`;
const FROM_TAG = `{${SHEET_DRAWING_NS}}from`;
const TO_TAG = `{${SHEET_DRAWING_NS}}to`;
const POS_TAG = `{${SHEET_DRAWING_NS}}pos`;
const EXT_TAG = `{${SHEET_DRAWING_NS}}ext`;
const COL_TAG = `{${SHEET_DRAWING_NS}}col`;
const COLOFF_TAG = `{${SHEET_DRAWING_NS}}colOff`;
const ROW_TAG = `{${SHEET_DRAWING_NS}}row`;
const ROWOFF_TAG = `{${SHEET_DRAWING_NS}}rowOff`;
const GRAPHIC_FRAME_TAG = `{${SHEET_DRAWING_NS}}graphicFrame`;
const CLIENT_DATA_TAG = `{${SHEET_DRAWING_NS}}clientData`;

const A_GRAPHIC_TAG = '{http://schemas.openxmlformats.org/drawingml/2006/main}graphic';
const A_GRAPHIC_DATA_TAG = '{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData';
const C_CHART_TAG = '{http://schemas.openxmlformats.org/drawingml/2006/chart}chart';

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const escapeAttr = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

const parseIntChild = (parent: XmlNode, tag: string): number | undefined => {
  const child = findChild(parent, tag);
  if (!child) return undefined;
  const n = Number.parseInt(child.text ?? '', 10);
  return Number.isInteger(n) ? n : undefined;
};

const parseAnchorMarker = (node: XmlNode): AnchorMarker | undefined => {
  const col = parseIntChild(node, COL_TAG);
  const colOff = parseIntChild(node, COLOFF_TAG);
  const row = parseIntChild(node, ROW_TAG);
  const rowOff = parseIntChild(node, ROWOFF_TAG);
  if (col === undefined || row === undefined) return undefined;
  return { col, colOff: colOff ?? 0, row, rowOff: rowOff ?? 0 };
};

const parsePoint2D = (node: XmlNode): Point2D | undefined => {
  const x = node.attrs['x'];
  const y = node.attrs['y'];
  if (!x || !y) return undefined;
  const xn = Number.parseInt(x, 10);
  const yn = Number.parseInt(y, 10);
  if (!Number.isFinite(xn) || !Number.isFinite(yn)) return undefined;
  return { x: xn, y: yn };
};

const parsePositiveSize2D = (node: XmlNode): PositiveSize2D | undefined => {
  const cx = node.attrs['cx'];
  const cy = node.attrs['cy'];
  if (!cx || !cy) return undefined;
  const cxn = Number.parseInt(cx, 10);
  const cyn = Number.parseInt(cy, 10);
  if (!Number.isFinite(cxn) || !Number.isFinite(cyn)) return undefined;
  return { cx: cxn, cy: cyn };
};

const parseChartReference = (node: XmlNode): ChartReference | undefined => {
  const graphic = findChild(node, GRAPHIC_FRAME_TAG);
  if (!graphic) return undefined;
  const aGraphic = findChild(graphic, A_GRAPHIC_TAG);
  const aGraphicData = aGraphic ? findChild(aGraphic, A_GRAPHIC_DATA_TAG) : undefined;
  const chart = aGraphicData ? findChild(aGraphicData, C_CHART_TAG) : undefined;
  if (!chart) return undefined;
  const rId = chart.attrs[`{${REL_NS}}id`];
  return rId !== undefined ? { rId } : {};
};

const parseAnchor = (node: XmlNode): DrawingItem | undefined => {
  let anchor: DrawingAnchor | undefined;
  if (node.name === ABSOLUTE_ANCHOR_TAG) {
    const pos = findChild(node, POS_TAG);
    const ext = findChild(node, EXT_TAG);
    if (!pos || !ext) return undefined;
    const p = parsePoint2D(pos);
    const e = parsePositiveSize2D(ext);
    if (!p || !e) return undefined;
    anchor = { kind: 'absolute', pos: p, ext: e };
  } else if (node.name === ONE_CELL_ANCHOR_TAG) {
    const fromEl = findChild(node, FROM_TAG);
    const ext = findChild(node, EXT_TAG);
    const from = fromEl ? parseAnchorMarker(fromEl) : undefined;
    const e = ext ? parsePositiveSize2D(ext) : undefined;
    if (!from || !e) return undefined;
    anchor = { kind: 'oneCell', from, ext: e };
  } else if (node.name === TWO_CELL_ANCHOR_TAG) {
    const fromEl = findChild(node, FROM_TAG);
    const toEl = findChild(node, TO_TAG);
    const from = fromEl ? parseAnchorMarker(fromEl) : undefined;
    const to = toEl ? parseAnchorMarker(toEl) : undefined;
    if (!from || !to) return undefined;
    const editAsRaw = node.attrs['editAs'];
    const validEditAs = editAsRaw === 'twoCell' || editAsRaw === 'oneCell' || editAsRaw === 'absolute';
    anchor = {
      kind: 'twoCell',
      from,
      to,
      ...(validEditAs ? { editAs: editAsRaw as 'twoCell' | 'oneCell' | 'absolute' } : {}),
    };
  } else {
    return undefined;
  }

  // Detect content kind. Stage-1 only models charts; everything else is
  // tagged "unsupported" with the original child tag name.
  const chart = parseChartReference(node);
  if (chart) {
    return { anchor, content: { kind: 'chart', chart } };
  }
  // Find the first child that isn't a marker/pos/ext/clientData.
  const skip = new Set([FROM_TAG, TO_TAG, POS_TAG, EXT_TAG, CLIENT_DATA_TAG]);
  for (const child of node.children) {
    if (skip.has(child.name)) continue;
    return { anchor, content: { kind: 'unsupported', rawTag: child.name } };
  }
  // Bare anchor with no content — record as unsupported with empty tag.
  return { anchor, content: { kind: 'unsupported', rawTag: '' } };
};

/** Parse a `xl/drawings/drawingN.xml` payload into a Drawing object. */
export function parseDrawingXml(bytes: Uint8Array | string): Drawing {
  const root = parseXml(bytes);
  if (root.name !== WS_DRAWING_TAG) {
    throw new OpenXmlSchemaError(`parseDrawingXml: root is "${root.name}", expected wsDr`);
  }
  const items: DrawingItem[] = [];
  // Document order matters — anchors must round-trip in their source
  // sequence regardless of their kind.
  for (const child of root.children) {
    if (child.name !== ABSOLUTE_ANCHOR_TAG && child.name !== ONE_CELL_ANCHOR_TAG && child.name !== TWO_CELL_ANCHOR_TAG)
      continue;
    const item = parseAnchor(child);
    if (item) items.push(item);
  }
  return makeDrawing(items);
}

const serializeMarker = (tag: string, m: AnchorMarker): string =>
  `<xdr:${tag}><xdr:col>${m.col}</xdr:col><xdr:colOff>${m.colOff}</xdr:colOff><xdr:row>${m.row}</xdr:row><xdr:rowOff>${m.rowOff}</xdr:rowOff></xdr:${tag}>`;

const serializeChartGraphicFrame = (chart: ChartReference, anchorIdx: number): string => {
  const rId = chart.rId ?? '';
  return [
    '<xdr:graphicFrame>',
    `<xdr:nvGraphicFramePr><xdr:cNvPr id="${anchorIdx + 2}" name="Chart ${anchorIdx + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`,
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>',
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">',
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${REL_NS}" r:id="${escapeAttr(rId)}"/>`,
    '</a:graphicData></a:graphic>',
    '</xdr:graphicFrame>',
  ].join('');
};

const serializeAnchor = (item: DrawingItem, idx: number): string => {
  const a = item.anchor;
  let body = '';
  if (a.kind === 'absolute') {
    body = `<xdr:pos x="${a.pos.x}" y="${a.pos.y}"/><xdr:ext cx="${a.ext.cx}" cy="${a.ext.cy}"/>`;
  } else if (a.kind === 'oneCell') {
    body = `${serializeMarker('from', a.from)}<xdr:ext cx="${a.ext.cx}" cy="${a.ext.cy}"/>`;
  } else {
    body = `${serializeMarker('from', a.from)}${serializeMarker('to', a.to)}`;
  }
  let content = '';
  if (item.content.kind === 'chart') {
    content = serializeChartGraphicFrame(item.content.chart, idx);
  } else {
    // Unsupported content: emit a graphicFrame with an empty chart ref so
    // Excel doesn't choke. Re-emitting unknown content verbatim is the
    // job of a later iteration (we don't carry the original XmlNode tree
    // through the data model in stage-1).
    content = serializeChartGraphicFrame({}, idx);
  }
  const editAs = a.kind === 'twoCell' && a.editAs ? ` editAs="${a.editAs}"` : '';
  const tag = a.kind === 'absolute' ? 'absoluteAnchor' : a.kind === 'oneCell' ? 'oneCellAnchor' : 'twoCellAnchor';
  return `<xdr:${tag}${editAs}>${body}${content}<xdr:clientData/></xdr:${tag}>`;
};

/** Serialise a Drawing to its `xl/drawings/drawingN.xml` payload. */
export function drawingToBytes(drawing: Drawing): Uint8Array {
  return new TextEncoder().encode(serializeDrawing(drawing));
}

export function serializeDrawing(drawing: Drawing): string {
  const parts: string[] = [
    XML_HEADER,
    `<xdr:wsDr xmlns:xdr="${SHEET_DRAWING_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`,
  ];
  for (let i = 0; i < drawing.items.length; i++) {
    const item = drawing.items[i];
    if (item) parts.push(serializeAnchor(item, i));
  }
  parts.push('</xdr:wsDr>');
  return parts.join('');
}
