// Worksheet XML reader. Per docs/plan/05-read-write.md §5.1.
//
// **Stage 1**: DOM-based reader for `<sheetData>/<row>/<c>` covering
// the common cell-value shapes — number / shared-string / boolean /
// error / inline string / formula. SAX iterparse + dimension /
// sheetView / mergeCells / hyperlinks etc. land in later iterations
// of the loop. The function signature is stable so the SAX swap won't
// break downstream callers.

import {
  type Cell,
  type ExcelErrorCode,
  type FormulaKind,
  setArrayFormula,
  setFormula,
  setSharedFormula,
} from '../cell/cell';
import { translateFormula } from '../formula/translate';
import type { Relationships } from '../packaging/relationships';
import { findById } from '../packaging/relationships';
import { coordinateToTuple, tupleToCoordinate } from '../utils/coordinate';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { ERROR_CODES } from '../utils/inference';
import { REL_NS, SHEET_MAIN_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, findChildren, type XmlNode } from '../xml/tree';
import { parseRange } from './cell-range';
import type { ColumnDimension, RowDimension } from './dimensions';
import { makeColumnDimension, makeRowDimension } from './dimensions';
import type { Hyperlink } from './hyperlinks';
import { makeHyperlink } from './hyperlinks';
import type { Pane, PaneState, PaneType, Selection, SheetView, SheetViewMode } from './views';
import { makeSheetView } from './views';
import { makeWorksheet, setCell, type Worksheet } from './worksheet';

const WORKSHEET_TAG = `{${SHEET_MAIN_NS}}worksheet`;
const SHEETDATA_TAG = `{${SHEET_MAIN_NS}}sheetData`;
const ROW_TAG = `{${SHEET_MAIN_NS}}row`;
const C_TAG = `{${SHEET_MAIN_NS}}c`;
const V_TAG = `{${SHEET_MAIN_NS}}v`;
const F_TAG = `{${SHEET_MAIN_NS}}f`;
const IS_TAG = `{${SHEET_MAIN_NS}}is`;
const T_TAG = `{${SHEET_MAIN_NS}}t`;
const MERGE_CELLS_TAG = `{${SHEET_MAIN_NS}}mergeCells`;
const MERGE_CELL_TAG = `{${SHEET_MAIN_NS}}mergeCell`;
const SHEET_VIEWS_TAG = `{${SHEET_MAIN_NS}}sheetViews`;
const SHEET_VIEW_TAG = `{${SHEET_MAIN_NS}}sheetView`;
const PANE_TAG = `{${SHEET_MAIN_NS}}pane`;
const SELECTION_TAG = `{${SHEET_MAIN_NS}}selection`;
const COLS_TAG = `{${SHEET_MAIN_NS}}cols`;
const COL_TAG = `{${SHEET_MAIN_NS}}col`;
const SHEET_FORMAT_PR_TAG = `{${SHEET_MAIN_NS}}sheetFormatPr`;
const HYPERLINKS_TAG = `{${SHEET_MAIN_NS}}hyperlinks`;
const HYPERLINK_TAG = `{${SHEET_MAIN_NS}}hyperlink`;

/** Inputs the worksheet reader needs from the surrounding workbook context. */
export interface WorksheetReadContext {
  /** Resolved shared-strings table. Pass `[]` when no sst is present. */
  sharedStrings: ReadonlyArray<string>;
  /** This worksheet's `_rels/sheetN.xml.rels`. Used to resolve external hyperlink targets. */
  rels?: Relationships;
}

/** Per-worksheet state for shared-formula expansion. */
interface SharedFormulaCache {
  origin: string;
  formula: string;
}

/**
 * Parse a `xl/worksheets/sheetN.xml` payload into a fully-populated
 * Worksheet. The returned worksheet's `title` matches the `title`
 * argument; the XML doesn't carry the sheet name (it lives in
 * `workbook.xml`).
 */
export function parseWorksheetXml(bytes: Uint8Array | string, title: string, ctx: WorksheetReadContext): Worksheet {
  const root = parseXml(bytes);
  if (root.name !== WORKSHEET_TAG) {
    throw new OpenXmlSchemaError(`parseWorksheetXml: root is "${root.name}", expected worksheet`);
  }
  const ws = makeWorksheet(title);

  // <sheetFormatPr> defaults — recorded so dimension-less sheets still
  // reflect any non-default workbook-wide row height / column width.
  const sheetFormatEl = findChild(root, SHEET_FORMAT_PR_TAG);
  if (sheetFormatEl) {
    const defaultColumnWidth = parseFloatAttr(sheetFormatEl.attrs['defaultColWidth']);
    if (defaultColumnWidth !== undefined) ws.defaultColumnWidth = defaultColumnWidth;
    const defaultRowHeight = parseFloatAttr(sheetFormatEl.attrs['defaultRowHeight']);
    if (defaultRowHeight !== undefined) ws.defaultRowHeight = defaultRowHeight;
  }

  // <cols> column dimensions — preserve runs verbatim (one entry per <col>).
  const colsEl = findChild(root, COLS_TAG);
  if (colsEl) {
    for (const c of findChildren(colsEl, COL_TAG)) {
      const dim = parseColumnDimension(c);
      ws.columnDimensions.set(dim.min, dim);
    }
  }

  const sheetData = findChild(root, SHEETDATA_TAG);
  if (sheetData) {
    const sharedFormulas = new Map<number, SharedFormulaCache>();
    for (const rowNode of findChildren(sheetData, ROW_TAG)) {
      const rowIdx = parseRowIndex(rowNode);
      maybeRecordRowDimension(ws, rowNode, rowIdx);
      let nextCol = 1;
      for (const cNode of findChildren(rowNode, C_TAG)) {
        const coord = parseCellCoord(cNode, rowIdx, nextCol);
        readCell(ws, cNode, coord, ctx, sharedFormulas);
        nextCol = coord.col + 1;
      }
    }
  }

  // <mergeCells> sits as a sibling of <sheetData>; pull each <mergeCell ref="…"/>
  // straight onto the worksheet's mergedCells list. We don't go through
  // mergeCells() here because its overlap check is for *new* merges; the
  // source xlsx is assumed valid.
  const mergeCellsEl = findChild(root, MERGE_CELLS_TAG);
  if (mergeCellsEl) {
    for (const m of findChildren(mergeCellsEl, MERGE_CELL_TAG)) {
      const ref = m.attrs['ref'];
      if (!ref) throw new OpenXmlSchemaError('worksheet: <mergeCell> missing @ref');
      ws.mergedCells.push(parseRange(ref));
    }
  }

  // <sheetViews> sits early in the worksheet but our load order doesn't
  // matter — reading it after sheetData keeps the loops separable.
  const sheetViewsEl = findChild(root, SHEET_VIEWS_TAG);
  if (sheetViewsEl) {
    for (const v of findChildren(sheetViewsEl, SHEET_VIEW_TAG)) {
      ws.views.push(parseSheetView(v));
    }
  }

  // <hyperlinks> — relations to external URLs come from the worksheet rels.
  const hyperlinksEl = findChild(root, HYPERLINKS_TAG);
  if (hyperlinksEl) {
    for (const h of findChildren(hyperlinksEl, HYPERLINK_TAG)) {
      ws.hyperlinks.push(parseHyperlink(h, ctx.rels));
    }
  }
  return ws;
}

const PANE_TYPES: ReadonlyArray<PaneType> = ['bottomRight', 'topRight', 'bottomLeft', 'topLeft'];
const PANE_STATES: ReadonlyArray<PaneState> = ['split', 'frozen', 'frozenSplit'];
const SHEET_VIEW_MODES: ReadonlyArray<SheetViewMode> = ['normal', 'pageBreakPreview', 'pageLayout'];

const parseFloatAttr = (raw: string | undefined): number | undefined => {
  if (raw === undefined) return undefined;
  const n = Number.parseFloat(raw);
  return Number.isFinite(n) ? n : undefined;
};

const parseIntegerAttr = (raw: string | undefined): number | undefined => {
  if (raw === undefined) return undefined;
  const n = Number.parseInt(raw, 10);
  return Number.isInteger(n) ? n : undefined;
};

const parseBoolXmlAttr = (raw: string | undefined): boolean | undefined => {
  if (raw === undefined) return undefined;
  if (raw === '1' || raw === 'true') return true;
  if (raw === '0' || raw === 'false') return false;
  return undefined;
};

const parseSheetView = (node: XmlNode): SheetView => {
  const opts: Partial<SheetView> = {
    workbookViewId: parseIntegerAttr(node.attrs['workbookViewId']) ?? 0,
  };
  const tabSelected = parseBoolXmlAttr(node.attrs['tabSelected']);
  if (tabSelected !== undefined) opts.tabSelected = tabSelected;
  const showGridLines = parseBoolXmlAttr(node.attrs['showGridLines']);
  if (showGridLines !== undefined) opts.showGridLines = showGridLines;
  const showRowColHeaders = parseBoolXmlAttr(node.attrs['showRowColHeaders']);
  if (showRowColHeaders !== undefined) opts.showRowColHeaders = showRowColHeaders;
  const showFormulas = parseBoolXmlAttr(node.attrs['showFormulas']);
  if (showFormulas !== undefined) opts.showFormulas = showFormulas;
  const showZeros = parseBoolXmlAttr(node.attrs['showZeros']);
  if (showZeros !== undefined) opts.showZeros = showZeros;
  const rightToLeft = parseBoolXmlAttr(node.attrs['rightToLeft']);
  if (rightToLeft !== undefined) opts.rightToLeft = rightToLeft;
  const view = node.attrs['view'];
  if (view !== undefined && (SHEET_VIEW_MODES as ReadonlyArray<string>).includes(view)) {
    opts.view = view as SheetViewMode;
  }
  if (node.attrs['topLeftCell']) opts.topLeftCell = node.attrs['topLeftCell'];
  const zoomScale = parseIntegerAttr(node.attrs['zoomScale']);
  if (zoomScale !== undefined) opts.zoomScale = zoomScale;
  const zoomScaleNormal = parseIntegerAttr(node.attrs['zoomScaleNormal']);
  if (zoomScaleNormal !== undefined) opts.zoomScaleNormal = zoomScaleNormal;

  const paneEl = findChild(node, PANE_TAG);
  if (paneEl) opts.pane = parsePane(paneEl);
  const selectionEl = findChild(node, SELECTION_TAG);
  if (selectionEl) opts.selection = parseSelection(selectionEl);
  return makeSheetView(opts);
};

const parsePane = (node: XmlNode): Pane => {
  const stateRaw = node.attrs['state'];
  const state: PaneState =
    stateRaw && (PANE_STATES as ReadonlyArray<string>).includes(stateRaw) ? (stateRaw as PaneState) : 'split';
  const pane: Pane = { state };
  const xSplit = parseFloatAttr(node.attrs['xSplit']);
  if (xSplit !== undefined) pane.xSplit = xSplit;
  const ySplit = parseFloatAttr(node.attrs['ySplit']);
  if (ySplit !== undefined) pane.ySplit = ySplit;
  if (node.attrs['topLeftCell']) pane.topLeftCell = node.attrs['topLeftCell'];
  const activePaneRaw = node.attrs['activePane'];
  if (activePaneRaw && (PANE_TYPES as ReadonlyArray<string>).includes(activePaneRaw)) {
    pane.activePane = activePaneRaw as PaneType;
  }
  return pane;
};

const parseSelection = (node: XmlNode): Selection => {
  const sel: Selection = {};
  const paneRaw = node.attrs['pane'];
  if (paneRaw && (PANE_TYPES as ReadonlyArray<string>).includes(paneRaw)) sel.pane = paneRaw as PaneType;
  if (node.attrs['activeCell']) sel.activeCell = node.attrs['activeCell'];
  if (node.attrs['sqref']) sel.sqref = node.attrs['sqref'];
  return sel;
};

const parseRowIndex = (rowNode: XmlNode): number => {
  const rAttr = rowNode.attrs['r'];
  if (rAttr === undefined) {
    throw new OpenXmlSchemaError('worksheet: <row> missing required @r');
  }
  const r = Number.parseInt(rAttr, 10);
  if (!Number.isInteger(r) || r < 1) {
    throw new OpenXmlSchemaError(`worksheet: <row r="${rAttr}"> is not a positive integer`);
  }
  return r;
};

const parseCellCoord = (cNode: XmlNode, rowIdx: number, fallbackCol: number): { row: number; col: number } => {
  const rAttr = cNode.attrs['r'];
  if (rAttr === undefined) {
    // Cells without @r take the next column slot in the current row.
    return { row: rowIdx, col: fallbackCol };
  }
  const t = coordinateToTuple(rAttr);
  if (t.row !== rowIdx) {
    throw new OpenXmlSchemaError(`worksheet: <c r="${rAttr}"> row ${t.row} disagrees with <row r="${rowIdx}">`);
  }
  return { row: t.row, col: t.col };
};

const readCell = (
  ws: Worksheet,
  cNode: XmlNode,
  coord: { row: number; col: number },
  ctx: WorksheetReadContext,
  sharedFormulas: Map<number, SharedFormulaCache>,
): void => {
  const t = cNode.attrs['t'] ?? 'n';
  const styleAttr = cNode.attrs['s'];
  const styleId = styleAttr === undefined ? 0 : parseStyleId(styleAttr);

  const fNode = findChild(cNode, F_TAG);
  const vNode = findChild(cNode, V_TAG);
  const isNode = findChild(cNode, IS_TAG);

  // ---- formula path ------------------------------------------------------
  if (fNode) {
    const cell = setCell(ws, coord.row, coord.col, null, styleId);
    const cachedRaw = vNode?.text;
    const cached = decodeCachedValue(cachedRaw, t);
    handleFormula(cell, fNode, coord, cached, sharedFormulas);
    return;
  }

  // ---- non-formula values ------------------------------------------------
  let value: number | string | boolean | { kind: 'error'; code: ExcelErrorCode } | null = null;
  switch (t) {
    case 'n':
      value = vNode?.text !== undefined && vNode.text !== '' ? Number.parseFloat(vNode.text) : null;
      break;
    case 's': {
      if (vNode?.text === undefined) {
        throw new OpenXmlSchemaError('worksheet: <c t="s"> missing <v>');
      }
      const idx = Number.parseInt(vNode.text, 10);
      if (!Number.isInteger(idx) || idx < 0) {
        throw new OpenXmlSchemaError(`worksheet: <c t="s"><v>${vNode.text}</v> is not a valid index`);
      }
      const sst = ctx.sharedStrings[idx];
      if (sst === undefined) {
        throw new OpenXmlSchemaError(
          `worksheet: shared-string index ${idx} out of range [0, ${ctx.sharedStrings.length})`,
        );
      }
      value = sst;
      break;
    }
    case 'b':
      value = vNode?.text === '1';
      break;
    case 'e': {
      const code = vNode?.text;
      if (code === undefined || !ERROR_CODES.has(code)) {
        throw new OpenXmlSchemaError(`worksheet: unknown error code "${code}" in <c t="e">`);
      }
      value = { kind: 'error', code: code as ExcelErrorCode };
      break;
    }
    case 'str':
      value = vNode?.text ?? '';
      break;
    case 'inlineStr':
      value = readInlineString(isNode);
      break;
    default:
      throw new OpenXmlSchemaError(`worksheet: unknown cell type t="${t}"`);
  }
  setCell(ws, coord.row, coord.col, value, styleId);
};

const parseStyleId = (raw: string): number => {
  const n = Number.parseInt(raw, 10);
  if (!Number.isInteger(n) || n < 0) {
    throw new OpenXmlSchemaError(`worksheet: <c s="${raw}"> is not a non-negative integer`);
  }
  return n;
};

/** Inline-string body — concatenates `<is>/<t>` text runs. Rich runs lose formatting in stage-1. */
const readInlineString = (isNode: XmlNode | undefined): string => {
  if (!isNode) return '';
  const direct = findChild(isNode, T_TAG);
  if (direct) return direct.text ?? '';
  // Rich inline string — concatenate every <r>/<t>.
  let out = '';
  for (const child of isNode.children) {
    const t = findChild(child, T_TAG);
    if (t?.text) out += t.text;
  }
  return out;
};

const decodeCachedValue = (raw: string | undefined, t: string): number | string | boolean | undefined => {
  if (raw === undefined || raw === '') return undefined;
  switch (t) {
    case 'n':
      return Number.parseFloat(raw);
    case 'b':
      return raw === '1';
    case 'str':
      return raw;
    case 'e':
      return raw;
    case 's':
      // Cached value of a formula resolving to a shared string is rare; keep as-is.
      return raw;
    default:
      return raw;
  }
};

const handleFormula = (
  cell: Cell,
  fNode: XmlNode,
  coord: { row: number; col: number },
  cached: number | string | boolean | undefined,
  sharedFormulas: Map<number, SharedFormulaCache>,
): void => {
  const tAttr = fNode.attrs['t'] ?? 'normal';
  const formula = fNode.text ?? '';
  const opts = cached !== undefined ? { cachedValue: cached } : undefined;
  switch (tAttr as FormulaKind) {
    case 'normal':
      setFormula(cell, formula, opts);
      return;
    case 'array': {
      const ref = fNode.attrs['ref'];
      if (!ref) {
        throw new OpenXmlSchemaError('worksheet: <f t="array"> missing @ref');
      }
      setArrayFormula(cell, ref, formula, opts);
      return;
    }
    case 'shared': {
      const siRaw = fNode.attrs['si'];
      if (siRaw === undefined) {
        throw new OpenXmlSchemaError('worksheet: <f t="shared"> missing @si');
      }
      const si = Number.parseInt(siRaw, 10);
      if (!Number.isInteger(si) || si < 0) {
        throw new OpenXmlSchemaError(`worksheet: <f t="shared" si="${siRaw}"> is not a valid index`);
      }
      const ref = fNode.attrs['ref'];
      if (formula.length > 0) {
        // Origin / first occurrence of the shared formula. The ref is required
        // by Excel but we accept its absence for resilience.
        sharedFormulas.set(si, { origin: tupleToCoordinate(coord.col, coord.row), formula });
        setSharedFormula(cell, si, formula, ref, opts);
        return;
      }
      // Subsequent reference — translate the cached origin formula.
      const cache = sharedFormulas.get(si);
      if (!cache) {
        throw new OpenXmlSchemaError(`worksheet: <f t="shared" si="${si}"/> with no preceding origin formula`);
      }
      const dest = tupleToCoordinate(coord.col, coord.row);
      // OOXML shared-formula text omits the leading '='; the translator
      // treats unprefixed input as a LITERAL and skips ref shifting, so we
      // re-prefix before translating and strip again on the way out.
      const translated = translateFormula(`=${cache.formula}`, cache.origin, { dest });
      const stripped = translated.startsWith('=') ? translated.slice(1) : translated;
      setSharedFormula(cell, si, stripped, undefined, opts);
      return;
    }
    case 'dataTable':
      // dataTable formulas are read in §5.5 (deferred). Drop the formula but
      // preserve the cached value so cells aren't lost on round-trip.
      setFormula(cell, formula, opts);
      return;
    default:
      throw new OpenXmlSchemaError(`worksheet: <f t="${tAttr}"> unknown formula kind`);
  }
};

const parseColumnDimension = (node: XmlNode): ColumnDimension => {
  const minRaw = node.attrs['min'];
  const maxRaw = node.attrs['max'];
  if (minRaw === undefined || maxRaw === undefined) {
    throw new OpenXmlSchemaError('worksheet: <col> missing required @min/@max');
  }
  const min = Number.parseInt(minRaw, 10);
  const max = Number.parseInt(maxRaw, 10);
  if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max < min) {
    throw new OpenXmlSchemaError(`worksheet: <col min="${minRaw}" max="${maxRaw}"> not a valid column run`);
  }
  const opts: Partial<Omit<ColumnDimension, 'min' | 'max'>> = {};
  const width = parseFloatAttr(node.attrs['width']);
  if (width !== undefined) opts.width = width;
  const customWidth = parseBoolXmlAttr(node.attrs['customWidth']);
  if (customWidth !== undefined) opts.customWidth = customWidth;
  const hidden = parseBoolXmlAttr(node.attrs['hidden']);
  if (hidden !== undefined) opts.hidden = hidden;
  const bestFit = parseBoolXmlAttr(node.attrs['bestFit']);
  if (bestFit !== undefined) opts.bestFit = bestFit;
  const outlineLevel = parseIntegerAttr(node.attrs['outlineLevel']);
  if (outlineLevel !== undefined) opts.outlineLevel = outlineLevel;
  const style = parseIntegerAttr(node.attrs['style']);
  if (style !== undefined) opts.style = style;
  const collapsed = parseBoolXmlAttr(node.attrs['collapsed']);
  if (collapsed !== undefined) opts.collapsed = collapsed;
  // makeColumnDimension fills min=col=max; rebuild with both range ends.
  return { ...makeColumnDimension(min, opts), max };
};

const parseHyperlink = (node: XmlNode, rels: Relationships | undefined): Hyperlink => {
  const ref = node.attrs['ref'];
  if (!ref) throw new OpenXmlSchemaError('worksheet: <hyperlink> missing @ref');
  const ridAttr = node.attrs[`{${REL_NS}}id`];
  const opts: Partial<Hyperlink> = { ref };
  if (ridAttr) {
    opts.rId = ridAttr;
    const rel = rels ? findById(rels, ridAttr) : undefined;
    if (rel) opts.target = rel.target;
  }
  if (node.attrs['location']) opts.location = node.attrs['location'];
  if (node.attrs['tooltip']) opts.tooltip = node.attrs['tooltip'];
  if (node.attrs['display']) opts.display = node.attrs['display'];
  return makeHyperlink(opts as Hyperlink);
};

const maybeRecordRowDimension = (ws: Worksheet, node: XmlNode, rowIdx: number): void => {
  const opts: Partial<RowDimension> = {};
  const ht = parseFloatAttr(node.attrs['ht']);
  if (ht !== undefined) opts.height = ht;
  const customHeight = parseBoolXmlAttr(node.attrs['customHeight']);
  if (customHeight !== undefined) opts.customHeight = customHeight;
  const hidden = parseBoolXmlAttr(node.attrs['hidden']);
  if (hidden !== undefined) opts.hidden = hidden;
  const outlineLevel = parseIntegerAttr(node.attrs['outlineLevel']);
  if (outlineLevel !== undefined) opts.outlineLevel = outlineLevel;
  const collapsed = parseBoolXmlAttr(node.attrs['collapsed']);
  if (collapsed !== undefined) opts.collapsed = collapsed;
  const style = parseIntegerAttr(node.attrs['s']);
  if (style !== undefined) opts.style = style;
  if (Object.keys(opts).length === 0) return;
  ws.rowDimensions.set(rowIdx, makeRowDimension(opts));
};
