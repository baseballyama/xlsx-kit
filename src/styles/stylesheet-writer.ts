// Stylesheet → xl/styles.xml writer. Per docs/plan/05-read-write.md §3.
//
// Pairs with parseStylesheetXml. The order of sections matches what
// Excel emits and what openpyxl writes — readers tolerate any order
// but Excel's diff-friendly layout helps when comparing fixtures.

import { toTree } from '../schema/serialize';
import { qname, SHEET_MAIN_NS } from '../xml/namespaces';
import { serializeXml } from '../xml/serializer';
import { el, type XmlNode } from '../xml/tree';
import { AlignmentSchema } from './alignment.schema';
import { BorderSchema } from './borders.schema';
import { fillToTree } from './fills.schema';
import { FontSchema } from './fonts.schema';
import { ProtectionSchema } from './protection.schema';
import type { CellXf, Stylesheet } from './stylesheet';

const STYLESHEET_TAG = qname(SHEET_MAIN_NS, 'styleSheet');
const FONTS_TAG = qname(SHEET_MAIN_NS, 'fonts');
const FILLS_TAG = qname(SHEET_MAIN_NS, 'fills');
const BORDERS_TAG = qname(SHEET_MAIN_NS, 'borders');
const NUMFMTS_TAG = qname(SHEET_MAIN_NS, 'numFmts');
const NUMFMT_TAG = qname(SHEET_MAIN_NS, 'numFmt');
const CELLSTYLEXFS_TAG = qname(SHEET_MAIN_NS, 'cellStyleXfs');
const CELLXFS_TAG = qname(SHEET_MAIN_NS, 'cellXfs');
const XF_TAG = qname(SHEET_MAIN_NS, 'xf');

/** Serialise a Stylesheet to its `xl/styles.xml` payload. */
export function stylesheetToBytes(ss: Stylesheet): Uint8Array {
  return serializeXml(buildStylesheetTree(ss));
}

function buildStylesheetTree(ss: Stylesheet): XmlNode {
  const root = el(STYLESHEET_TAG);

  // numFmts (custom only — built-ins are implicit).
  if (ss.numFmts.size > 0) {
    const numFmtsEl = el(NUMFMTS_TAG, { count: String(ss.numFmts.size) });
    const ids = [...ss.numFmts.keys()].sort((a, b) => a - b);
    for (const id of ids) {
      const code = ss.numFmts.get(id);
      if (code === undefined) continue;
      numFmtsEl.children.push(el(NUMFMT_TAG, { numFmtId: String(id), formatCode: code }));
    }
    root.children.push(numFmtsEl);
  }

  // fonts
  const fontsEl = el(FONTS_TAG, { count: String(ss.fonts.length) });
  for (const f of ss.fonts) fontsEl.children.push(toTree(f, FontSchema));
  root.children.push(fontsEl);

  // fills (use fillToTree — fillFromTree's symmetric writer)
  const fillsEl = el(FILLS_TAG, { count: String(ss.fills.length) });
  for (const fill of ss.fills) fillsEl.children.push(fillToTree(fill));
  root.children.push(fillsEl);

  // borders
  const bordersEl = el(BORDERS_TAG, { count: String(ss.borders.length) });
  for (const b of ss.borders) bordersEl.children.push(toTree(b, BorderSchema));
  root.children.push(bordersEl);

  // cellStyleXfs (always emitted — Excel rejects styles.xml that omits it
  // when cellXfs reference an xfId).
  const cellStyleXfsEl = el(CELLSTYLEXFS_TAG, {
    count: String(Math.max(ss.cellStyleXfs.length, 1)),
  });
  if (ss.cellStyleXfs.length === 0) {
    cellStyleXfsEl.children.push(el(XF_TAG, { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0' }));
  } else {
    for (const xf of ss.cellStyleXfs) cellStyleXfsEl.children.push(cellXfToTree(xf));
  }
  root.children.push(cellStyleXfsEl);

  // cellXfs — same fallback for empty pools.
  const cellXfsEl = el(CELLXFS_TAG, {
    count: String(Math.max(ss.cellXfs.length, 1)),
  });
  if (ss.cellXfs.length === 0) {
    cellXfsEl.children.push(el(XF_TAG, { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0', xfId: '0' }));
  } else {
    for (const xf of ss.cellXfs) cellXfsEl.children.push(cellXfToTree(xf));
  }
  root.children.push(cellXfsEl);

  return root;
}

const cellXfToTree = (xf: CellXf): XmlNode => {
  const attrs: Record<string, string> = {
    numFmtId: String(xf.numFmtId),
    fontId: String(xf.fontId),
    fillId: String(xf.fillId),
    borderId: String(xf.borderId),
  };
  if (xf.xfId !== undefined) attrs['xfId'] = String(xf.xfId);
  if (xf.applyFont) attrs['applyFont'] = '1';
  if (xf.applyFill) attrs['applyFill'] = '1';
  if (xf.applyBorder) attrs['applyBorder'] = '1';
  if (xf.applyNumberFormat) attrs['applyNumberFormat'] = '1';
  if (xf.applyAlignment) attrs['applyAlignment'] = '1';
  if (xf.applyProtection) attrs['applyProtection'] = '1';
  if (xf.pivotButton) attrs['pivotButton'] = '1';
  if (xf.quotePrefix) attrs['quotePrefix'] = '1';
  const node = el(XF_TAG, attrs);
  if (xf.alignment) node.children.push(toTree(xf.alignment, AlignmentSchema));
  if (xf.protection) node.children.push(toTree(xf.protection, ProtectionSchema));
  return node;
};
