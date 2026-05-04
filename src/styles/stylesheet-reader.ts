// xl/styles.xml → Stylesheet reader. Per docs/plan/05-read-write.md §3.
//
// The reader preserves slot ordering exactly — cellXfs[3] in the source
// xlsx must come back as cellXfs[3] in the loaded Workbook because every
// `<c s="3">` reference depends on the index, not on the value's identity.
// That rules out the `addFont` / `addCellXf` dedup helpers (which collapse
// equal values); instead we push raw entries and rebuild the
// `_*IdByKey` maps at the end so subsequent edits go back through dedup.

import { fromTree } from '../schema/serialize';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { stableStringify } from '../utils/stable-stringify';
import { qname, SHEET_MAIN_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, findChildren, type XmlNode } from '../xml/tree';
import { AlignmentSchema } from './alignment.schema';
import type { Border } from './borders';
import { BorderSchema } from './borders.schema';
import type { Fill } from './fills';
import { fillFromTree } from './fills.schema';
import type { Font } from './fonts';
import { FontSchema } from './fonts.schema';
import type { NumberFormat } from './numbers';
import { ProtectionSchema } from './protection.schema';
import { type CellXf, makeStylesheet, type Stylesheet } from './stylesheet';

const STYLESHEET_TAG = qname(SHEET_MAIN_NS, 'styleSheet');
const FONTS_TAG = qname(SHEET_MAIN_NS, 'fonts');
const FILLS_TAG = qname(SHEET_MAIN_NS, 'fills');
const BORDERS_TAG = qname(SHEET_MAIN_NS, 'borders');
const NUMFMTS_TAG = qname(SHEET_MAIN_NS, 'numFmts');
const NUMFMT_TAG = qname(SHEET_MAIN_NS, 'numFmt');
const CELLXFS_TAG = qname(SHEET_MAIN_NS, 'cellXfs');
const CELLSTYLEXFS_TAG = qname(SHEET_MAIN_NS, 'cellStyleXfs');
const XF_TAG = qname(SHEET_MAIN_NS, 'xf');
const FONT_TAG = qname(SHEET_MAIN_NS, 'font');
const FILL_TAG = qname(SHEET_MAIN_NS, 'fill');
const BORDER_TAG = qname(SHEET_MAIN_NS, 'border');
const ALIGNMENT_TAG = qname(SHEET_MAIN_NS, 'alignment');
const PROTECTION_TAG = qname(SHEET_MAIN_NS, 'protection');

/**
 * Parse `xl/styles.xml` and return a fully-populated {@link Stylesheet}.
 * Slot ordering is preserved verbatim; the dedup index Maps are rebuilt
 * after the fact so future `addFont` / `addCellXf` calls keep working.
 */
export function parseStylesheetXml(bytes: Uint8Array | string): Stylesheet {
  const root = parseXml(bytes);
  if (root.name !== STYLESHEET_TAG) {
    throw new OpenXmlSchemaError(`parseStylesheetXml: root is "${root.name}", expected styleSheet`);
  }

  // Start from a Stylesheet whose default-pool entries we'll wholesale replace
  // with the XML's actual contents.
  const ss = makeStylesheet();
  ss.fonts.length = 0;
  ss.fills.length = 0;
  ss.borders.length = 0;

  for (const fontEl of findInSection(root, FONTS_TAG, FONT_TAG)) {
    ss.fonts.push(fromTree(fontEl, FontSchema));
  }
  for (const fillEl of findInSection(root, FILLS_TAG, FILL_TAG)) {
    ss.fills.push(fillFromTree(fillEl));
  }
  for (const borderEl of findInSection(root, BORDERS_TAG, BORDER_TAG)) {
    ss.borders.push(fromTree(borderEl, BorderSchema));
  }

  // numFmts is a Map<id, code>; id is on the element so a sparse Map works.
  for (const numFmtEl of findInSection(root, NUMFMTS_TAG, NUMFMT_TAG)) {
    const nf = parseNumFmt(numFmtEl);
    ss.numFmts.set(nf.numFmtId, nf.formatCode);
  }

  for (const xfEl of findInSection(root, CELLSTYLEXFS_TAG, XF_TAG)) {
    ss.cellStyleXfs.push(parseCellXf(xfEl));
  }
  for (const xfEl of findInSection(root, CELLXFS_TAG, XF_TAG)) {
    ss.cellXfs.push(parseCellXf(xfEl));
  }

  rebuildIndexes(ss);
  return ss;
}

/** Drill from `<styleSheet>` into a numbered section and yield every matching child. */
function findInSection(root: XmlNode, sectionTag: string, itemTag: string): XmlNode[] {
  const section = findChild(root, sectionTag);
  if (!section) return [];
  return findChildren(section, itemTag);
}

const parseNumFmt = (node: XmlNode): NumberFormat => {
  const idAttr = node.attrs['numFmtId'];
  const code = node.attrs['formatCode'];
  if (idAttr === undefined) {
    throw new OpenXmlSchemaError('styles: <numFmt> missing @numFmtId');
  }
  if (code === undefined) {
    throw new OpenXmlSchemaError(`styles: <numFmt numFmtId="${idAttr}"> missing @formatCode`);
  }
  const numFmtId = Number.parseInt(idAttr, 10);
  if (!Number.isInteger(numFmtId) || numFmtId < 0) {
    throw new OpenXmlSchemaError(`styles: <numFmt numFmtId="${idAttr}"> is not a non-negative integer`);
  }
  return { numFmtId, formatCode: code };
};

const parseIntAttr = (raw: string | undefined, label: string): number | undefined => {
  if (raw === undefined) return undefined;
  const n = Number.parseInt(raw, 10);
  if (!Number.isInteger(n) || n < 0) {
    throw new OpenXmlSchemaError(`styles: <xf ${label}="${raw}"> is not a non-negative integer`);
  }
  return n;
};

const parseBoolAttr = (raw: string | undefined): boolean | undefined => {
  if (raw === undefined) return undefined;
  if (raw === '1' || raw === 'true') return true;
  if (raw === '0' || raw === 'false') return false;
  return undefined;
};

const parseCellXf = (node: XmlNode): CellXf => {
  const fontId = parseIntAttr(node.attrs['fontId'], 'fontId') ?? 0;
  const fillId = parseIntAttr(node.attrs['fillId'], 'fillId') ?? 0;
  const borderId = parseIntAttr(node.attrs['borderId'], 'borderId') ?? 0;
  const numFmtId = parseIntAttr(node.attrs['numFmtId'], 'numFmtId') ?? 0;
  const xfId = parseIntAttr(node.attrs['xfId'], 'xfId');

  const applyFont = parseBoolAttr(node.attrs['applyFont']);
  const applyFill = parseBoolAttr(node.attrs['applyFill']);
  const applyBorder = parseBoolAttr(node.attrs['applyBorder']);
  const applyNumberFormat = parseBoolAttr(node.attrs['applyNumberFormat']);
  const applyAlignment = parseBoolAttr(node.attrs['applyAlignment']);
  const applyProtection = parseBoolAttr(node.attrs['applyProtection']);
  const pivotButton = parseBoolAttr(node.attrs['pivotButton']);
  const quotePrefix = parseBoolAttr(node.attrs['quotePrefix']);

  const alignmentEl = findChild(node, ALIGNMENT_TAG);
  const protectionEl = findChild(node, PROTECTION_TAG);

  return {
    fontId,
    fillId,
    borderId,
    numFmtId,
    ...(xfId !== undefined ? { xfId } : {}),
    ...(alignmentEl ? { alignment: fromTree(alignmentEl, AlignmentSchema) } : {}),
    ...(protectionEl ? { protection: fromTree(protectionEl, ProtectionSchema) } : {}),
    ...(applyFont !== undefined ? { applyFont } : {}),
    ...(applyFill !== undefined ? { applyFill } : {}),
    ...(applyBorder !== undefined ? { applyBorder } : {}),
    ...(applyNumberFormat !== undefined ? { applyNumberFormat } : {}),
    ...(applyAlignment !== undefined ? { applyAlignment } : {}),
    ...(applyProtection !== undefined ? { applyProtection } : {}),
    ...(pivotButton !== undefined ? { pivotButton } : {}),
    ...(quotePrefix !== undefined ? { quotePrefix } : {}),
  };
};

/** Repopulate the `_*IdByKey` maps from the freshly-loaded pool entries. */
function rebuildIndexes(ss: Stylesheet): void {
  ss._fontIdByKey = buildKeyIndex<Font>(ss.fonts);
  ss._fillIdByKey = buildKeyIndex<Fill>(ss.fills);
  ss._borderIdByKey = buildKeyIndex<Border>(ss.borders);
  ss._xfIdByKey = buildKeyIndex<CellXf>(ss.cellXfs);
  ss._styleXfIdByKey = buildKeyIndex<CellXf>(ss.cellStyleXfs);
  ss._numFmtIdByCode = new Map();
  for (const [id, code] of ss.numFmts) ss._numFmtIdByCode.set(code, id);
}

const buildKeyIndex = <T>(arr: ReadonlyArray<T>): Map<string, number> => {
  const m = new Map<string, number>();
  for (let i = 0; i < arr.length; i++) {
    const key = stableStringify(arr[i] as T);
    if (!m.has(key)) m.set(key, i);
  }
  return m;
};
