// Worksheet XML writer. Per docs/plan/05-read-write.md §5.2.
//
// **Stage 1**: hand-rolled serialiser for `<sheetData>/<row>/<c>`
// covering the common cell-value shapes — number / string (via shared
// strings) / boolean / error / formula. Streaming through
// XmlStreamWriter + dimension / sheetView / cols / mergeCells lands in
// later iterations of the loop.
//
// The acceptance criterion (1M cell write in ~5s on M1) needs SAX, but
// stage-1 prioritises correctness — once loadWorkbook → saveWorkbook
// round-trips, we can swap the body for a streaming writer without
// callers noticing.

import { type Cell, type CellValue, type ExcelErrorCode, type FormulaValue, getCoordinate } from '../cell/cell';
import type { Relationships } from '../packaging/relationships';
import { dateToExcel, durationToExcel } from '../utils/datetime';
import { escapeCellString } from '../utils/escape';
import { OpenXmlSchemaError } from '../utils/exceptions';
import type { SharedStringsTable } from '../workbook/shared-strings';
import { addSharedString } from '../workbook/shared-strings';
import { SHEET_MAIN_NS } from '../xml/namespaces';
import { serializeXml } from '../xml/serializer';
import type { XmlNode } from '../xml/tree';
import type { AutoFilter } from './auto-filter';
import { multiCellRangeToString, rangeToString } from './cell-range';
import type { ConditionalFormatting, ConditionalFormattingRule } from './conditional-formatting';
import type { DataValidation } from './data-validations';
import type { ColumnDimension, RowDimension } from './dimensions';
import type { CellWatch, IgnoredError } from './errors';
import type { Hyperlink } from './hyperlinks';
import type { SheetProperties } from './properties';
import type { SheetProtection } from './protection';
import type { Pane, Selection, SheetView } from './views';
import type { Worksheet } from './worksheet';

const HYPERLINK_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

export interface WorksheetWriteContext {
  /** Accumulator the writer mutates as it emits string cells. */
  sharedStrings: SharedStringsTable;
  /**
   * Workbook epoch for `Date` / `{kind:'duration'}` cell serialisation.
   * `true` = Mac 1904 epoch; `false` (default) = Windows 1900 epoch.
   * Modern Excel emits 1900-based serials; 1904 only appears in legacy
   * Mac workbooks but the round-trip respects the workbook setting.
   */
  date1904?: boolean;
  /**
   * Worksheet rels collector. The writer pushes a rel per external
   * hyperlink and per Table, allocating ids `rId1..rIdN`. Caller emits
   * the resulting relationships file alongside the worksheet part.
   */
  rels?: Relationships;
  /**
   * Table rel allocator. saveWorkbook hands in a callback that registers
   * a table part under a workbook-global tableN counter and returns the
   * relPath ("../tables/tableN.xml") + the worksheet-rels rId. Called
   * once per `ws.tables` entry while serialising.
   */
  registerTable?: (table: import('./table').TableDefinition) => { rId: string };
  /**
   * Comments / VML drawing allocator. saveWorkbook emits the comments
   * part + a placeholder VML drawing for all comments on the sheet, and
   * returns the worksheet-rels rId for the VML — which the writer
   * splats into `<legacyDrawing r:id>`. Called once per worksheet that
   * carries any comments.
   */
  registerComments?: (comments: ReadonlyArray<import('./comments').LegacyComment>) => { vmlRelId: string };
  /**
   * Drawing allocator. saveWorkbook emits xl/drawings/drawingN.xml under
   * a workbook-global counter, registers a `${REL_NS}/drawing` rel on
   * the worksheet rels, and returns the worksheet-rels rId — splatted
   * into `<drawing r:id>` by the writer. Called once when ws.drawing is
   * set.
   */
  registerDrawing?: (drawing: import('../drawing/drawing').Drawing) => { rId: string };
}

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Serialise a worksheet to its `xl/worksheets/sheetN.xml` payload. The
 * function mutates `ctx.sharedStrings` — every plain-string cell adds
 * (or dedupes) into the table; the caller is responsible for emitting
 * the resulting sst at the end of the package write.
 */
export function worksheetToBytes(ws: Worksheet, ctx: WorksheetWriteContext): Uint8Array {
  return new TextEncoder().encode(serializeWorksheet(ws, ctx));
}

export function serializeWorksheet(ws: Worksheet, ctx: WorksheetWriteContext): string {
  const parts: string[] = [
    XML_HEADER,
    `<worksheet xmlns="${SHEET_MAIN_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`,
  ];
  if (ws.sheetProperties) {
    const sp = serializeSheetProperties(ws.sheetProperties);
    if (sp) parts.push(sp);
  }
  if (ws.bodyExtras?.beforeSheetData) {
    for (const node of ws.bodyExtras.beforeSheetData) parts.push(serializeBodyExtraNode(node));
  }
  parts.push(serializeDimension(ws));
  if (ws.views.length > 0) parts.push(serializeSheetViews(ws.views));
  const sheetFormatPr = serializeSheetFormatPr(ws);
  if (sheetFormatPr) parts.push(sheetFormatPr);
  if (ws.columnDimensions.size > 0) parts.push(serializeCols(ws.columnDimensions));
  parts.push('<sheetData>');
  // Iterate rows in numeric order so writer output is deterministic.
  const rowKeys = [...ws.rows.keys()].sort((a, b) => a - b);
  // Walk the union of populated rows + rowDimension entries so dimension-
  // only rows (e.g. an empty row with custom height) still emit.
  const rowKeyUnion = new Set<number>(rowKeys);
  for (const k of ws.rowDimensions.keys()) rowKeyUnion.add(k);
  const sortedRowKeys = [...rowKeyUnion].sort((a, b) => a - b);
  for (const rowIdx of sortedRowKeys) {
    const row = ws.rows.get(rowIdx);
    const dim = ws.rowDimensions.get(rowIdx);
    if ((!row || row.size === 0) && !dim) continue;
    const dimAttrs = dim ? serializeRowDimensionAttrs(dim) : '';
    if (!row || row.size === 0) {
      parts.push(`<row r="${rowIdx}"${dimAttrs}/>`);
      continue;
    }
    const colKeys = [...row.keys()].sort((a, b) => a - b);
    parts.push(`<row r="${rowIdx}"${dimAttrs}>`);
    for (const colIdx of colKeys) {
      const cell = row.get(colIdx);
      if (cell) parts.push(serializeCell(cell, ctx));
    }
    parts.push('</row>');
  }
  parts.push('</sheetData>');
  // sheetProtection sits between sheetData and mergeCells per
  // ECMA-376 §18.3.1.85.
  if (ws.sheetProtection) {
    const sp = serializeSheetProtection(ws.sheetProtection);
    if (sp) parts.push(sp);
  }
  // Excel's element order: autoFilter sits between sheetData and mergeCells.
  if (ws.autoFilter) parts.push(serializeAutoFilter(ws.autoFilter));
  if (ws.mergedCells.length > 0) {
    parts.push(`<mergeCells count="${ws.mergedCells.length}">`);
    for (const range of ws.mergedCells) {
      parts.push(`<mergeCell ref="${rangeToString(range)}"/>`);
    }
    parts.push('</mergeCells>');
  }
  for (const cf of ws.conditionalFormatting) {
    parts.push(serializeConditionalFormatting(cf));
  }
  if (ws.dataValidations.length > 0) {
    parts.push(serializeDataValidations(ws.dataValidations));
  }
  if (ws.hyperlinks.length > 0) {
    parts.push(serializeHyperlinks(ws.hyperlinks, ctx.rels));
  }
  // afterSheetData extras live between hyperlinks and the drawing/
  // legacyDrawing/tableParts tail. ECMA-376 puts printOptions / pageMargins
  // / pageSetup / headerFooter / rowBreaks / colBreaks here; oleObjects /
  // controls / picture / legacyDrawingHF technically belong AFTER drawing
  // and BEFORE tableParts, but Excel reads them in either spot. extLst is
  // strictly the last child per ECMA, but landing it before tableParts
  // keeps round-trip Excel-compatible without requiring fine-grained
  // positional tracking.
  if (ws.bodyExtras?.afterSheetData) {
    for (const node of ws.bodyExtras.afterSheetData) parts.push(serializeBodyExtraNode(node));
  }
  if (ws.cellWatches.length > 0) {
    parts.push(serializeCellWatches(ws.cellWatches));
  }
  if (ws.ignoredErrors.length > 0) {
    parts.push(serializeIgnoredErrors(ws.ignoredErrors));
  }
  if (ws.drawing && ctx.registerDrawing) {
    const { rId } = ctx.registerDrawing(ws.drawing);
    parts.push(`<drawing r:id="${escapeXmlAttr(rId)}"/>`);
  }
  if (ws.legacyComments.length > 0 && ctx.registerComments) {
    const { vmlRelId } = ctx.registerComments(ws.legacyComments);
    parts.push(`<legacyDrawing r:id="${escapeXmlAttr(vmlRelId)}"/>`);
  }
  if (ws.tables.length > 0 && ctx.registerTable) {
    parts.push(`<tableParts count="${ws.tables.length}">`);
    for (const t of ws.tables) {
      const { rId } = ctx.registerTable(t);
      parts.push(`<tablePart r:id="${escapeXmlAttr(rId)}"/>`);
    }
    parts.push('</tableParts>');
  }
  parts.push('</worksheet>');
  return parts.join('');
}

const escapeXmlText = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const escapeXmlAttr = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

const serializeDimension = (ws: Worksheet): string => {
  let minRow = Infinity;
  let maxRow = -Infinity;
  let minCol = Infinity;
  let maxCol = -Infinity;
  for (const [rowIdx, row] of ws.rows) {
    if (row.size === 0) continue;
    for (const colIdx of row.keys()) {
      if (rowIdx < minRow) minRow = rowIdx;
      if (rowIdx > maxRow) maxRow = rowIdx;
      if (colIdx < minCol) minCol = colIdx;
      if (colIdx > maxCol) maxCol = colIdx;
    }
  }
  if (!Number.isFinite(minRow)) return '<dimension ref="A1"/>';
  // Compose the ref using a temporary Cell-like to share the column-letter logic.
  const col1 = colLetters(minCol);
  const col2 = colLetters(maxCol);
  const ref = minRow === maxRow && minCol === maxCol ? `${col1}${minRow}` : `${col1}${minRow}:${col2}${maxRow}`;
  return `<dimension ref="${ref}"/>`;
};

const colLetters = (n: number): string => {
  let m = n;
  let out = '';
  while (m > 0) {
    m -= 1;
    out = String.fromCharCode(65 + (m % 26)) + out;
    m = Math.floor(m / 26);
  }
  return out;
};

/**
 * Serialise a single cell into its `<c .../>` element. Exported so the
 * streaming write-only path can emit cells row-by-row without going
 * through the full Worksheet model — see src/streaming/write-only.ts.
 *
 * Only `ctx.sharedStrings` is consulted for plain-string cells; the
 * other context fields are used by the worksheet-level serializer that
 * wraps this helper.
 */
export const serializeCell = (cell: Cell, ctx: WorksheetWriteContext): string => {
  const ref = getCoordinate(cell);
  const styleAttr = cell.styleId === 0 ? '' : ` s="${cell.styleId}"`;
  const value = cell.value;

  if (value === null) {
    // Empty-but-styled cells still need to emit so styleId survives the round-trip.
    if (cell.styleId === 0) return '';
    return `<c r="${ref}"${styleAttr}/>`;
  }

  // Formula values come first since the discriminated union check is cheap.
  if (typeof value === 'object' && value !== null && (value as { kind?: string }).kind === 'formula') {
    return serializeFormulaCell(ref, styleAttr, value as FormulaValue);
  }
  if (typeof value === 'object' && value !== null && (value as { kind?: string }).kind === 'error') {
    const code = (value as { kind: 'error'; code: ExcelErrorCode }).code;
    return `<c r="${ref}"${styleAttr} t="e"><v>${code}</v></c>`;
  }
  if (typeof value === 'object' && value !== null && (value as { kind?: string }).kind === 'rich-text') {
    // Stage-1: flatten rich text into a plain string, dedup via sst.
    const runs = (value as { kind: 'rich-text'; runs: ReadonlyArray<{ text: string }> }).runs;
    let flat = '';
    for (const run of runs) flat += run.text;
    const id = addSharedString(ctx.sharedStrings, flat);
    return `<c r="${ref}"${styleAttr} t="s"><v>${id}</v></c>`;
  }

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new OpenXmlSchemaError(`worksheet: cannot serialise non-finite number at ${ref}`);
    }
    return `<c r="${ref}"${styleAttr}><v>${serializeNumber(value)}</v></c>`;
  }
  if (typeof value === 'boolean') {
    return `<c r="${ref}"${styleAttr} t="b"><v>${value ? '1' : '0'}</v></c>`;
  }
  if (typeof value === 'string') {
    const id = addSharedString(ctx.sharedStrings, value);
    return `<c r="${ref}"${styleAttr} t="s"><v>${id}</v></c>`;
  }
  if (value instanceof Date) {
    // Excel stores dates as serial numbers under the workbook epoch.
    // The cell shows as a date only if the styleId points at a number-
    // format with a date code — caller-managed; openpyxl behaves the
    // same way. We just emit the serial.
    const serial = dateToExcel(value, { epoch: ctx.date1904 ? 'mac' : 'windows' });
    return `<c r="${ref}"${styleAttr}><v>${serializeNumber(serial)}</v></c>`;
  }
  if (typeof value === 'object' && value !== null && (value as { kind?: string }).kind === 'duration') {
    const ms = (value as { kind: 'duration'; ms: number }).ms;
    const serial = durationToExcel(ms);
    return `<c r="${ref}"${styleAttr}><v>${serializeNumber(serial)}</v></c>`;
  }
  throw new OpenXmlSchemaError(`worksheet: unsupported cell value kind at ${ref}: ${describeValue(value)}`);
};

const serializeNumber = (n: number): string => {
  // Match Excel's round-trip preference: integers stay as integers, doubles
  // use the JS default representation. We don't need scientific-notation
  // tweaking here — Excel reads either form fine.
  return Number.isInteger(n) ? n.toFixed(0) : String(n);
};

const serializeFormulaCell = (ref: string, styleAttr: string, f: FormulaValue): string => {
  const fAttrs: string[] = [];
  if (f.t !== 'normal') fAttrs.push(`t="${f.t}"`);
  if (f.ref !== undefined) fAttrs.push(`ref="${escapeXmlAttr(f.ref)}"`);
  if (f.si !== undefined) fAttrs.push(`si="${f.si}"`);
  // Data-table formula attrs — only relevant when t === 'dataTable'.
  if (f.r1 !== undefined) fAttrs.push(`r1="${escapeXmlAttr(f.r1)}"`);
  if (f.r2 !== undefined) fAttrs.push(`r2="${escapeXmlAttr(f.r2)}"`);
  if (f.dt2D) fAttrs.push('dt2D="1"');
  if (f.dtr) fAttrs.push('dtr="1"');
  if (f.del1) fAttrs.push('del1="1"');
  if (f.del2) fAttrs.push('del2="1"');
  if (f.aca) fAttrs.push('aca="1"');
  if (f.ca) fAttrs.push('ca="1"');
  const fAttrStr = fAttrs.length > 0 ? ` ${fAttrs.join(' ')}` : '';
  const formulaText = escapeXmlText(escapeCellString(f.formula));
  const fEl = formulaText.length > 0 ? `<f${fAttrStr}>${formulaText}</f>` : `<f${fAttrStr}/>`;

  let valueAttr = '';
  let vEl = '';
  const cached = f.cachedValue;
  if (cached !== undefined) {
    if (typeof cached === 'number') {
      vEl = `<v>${serializeNumber(cached)}</v>`;
    } else if (typeof cached === 'boolean') {
      valueAttr = ' t="b"';
      vEl = `<v>${cached ? '1' : '0'}</v>`;
    } else {
      // String result of a formula — use t="str", not the sst path.
      valueAttr = ' t="str"';
      vEl = `<v>${escapeXmlText(escapeCellString(cached))}</v>`;
    }
  }
  return `<c r="${ref}"${styleAttr}${valueAttr}>${fEl}${vEl}</c>`;
};

const describeValue = (value: CellValue): string => {
  if (value === null) return 'null';
  if (typeof value === 'object' && value !== null) {
    const kind = (value as { kind?: string }).kind;
    return kind ? `{ kind: "${kind}" }` : 'object';
  }
  return typeof value;
};

const xmlBoolAttr = (key: string, v: boolean | undefined): string =>
  v === undefined ? '' : ` ${key}="${v ? '1' : '0'}"`;

const serializeSheetViews = (views: ReadonlyArray<SheetView>): string => {
  const parts: string[] = ['<sheetViews>'];
  for (const v of views) parts.push(serializeSheetView(v));
  parts.push('</sheetViews>');
  return parts.join('');
};

const serializeSheetView = (v: SheetView): string => {
  let attrs = '';
  attrs += xmlBoolAttr('tabSelected', v.tabSelected);
  attrs += xmlBoolAttr('showGridLines', v.showGridLines);
  attrs += xmlBoolAttr('showRowColHeaders', v.showRowColHeaders);
  attrs += xmlBoolAttr('showFormulas', v.showFormulas);
  attrs += xmlBoolAttr('showZeros', v.showZeros);
  attrs += xmlBoolAttr('rightToLeft', v.rightToLeft);
  if (v.view) attrs += ` view="${v.view}"`;
  if (v.topLeftCell) attrs += ` topLeftCell="${escapeXmlAttr(v.topLeftCell)}"`;
  if (v.zoomScale !== undefined) attrs += ` zoomScale="${v.zoomScale}"`;
  if (v.zoomScaleNormal !== undefined) attrs += ` zoomScaleNormal="${v.zoomScaleNormal}"`;
  attrs += ` workbookViewId="${v.workbookViewId}"`;

  const inner: string[] = [];
  if (v.pane) inner.push(serializePane(v.pane));
  if (v.selection) inner.push(serializeSelection(v.selection));

  if (inner.length === 0) return `<sheetView${attrs}/>`;
  return `<sheetView${attrs}>${inner.join('')}</sheetView>`;
};

const serializePane = (p: Pane): string => {
  let attrs = '';
  if (p.xSplit !== undefined) attrs += ` xSplit="${p.xSplit}"`;
  if (p.ySplit !== undefined) attrs += ` ySplit="${p.ySplit}"`;
  if (p.topLeftCell) attrs += ` topLeftCell="${escapeXmlAttr(p.topLeftCell)}"`;
  if (p.activePane) attrs += ` activePane="${p.activePane}"`;
  attrs += ` state="${p.state}"`;
  return `<pane${attrs}/>`;
};

const serializeSelection = (s: Selection): string => {
  let attrs = '';
  if (s.pane) attrs += ` pane="${s.pane}"`;
  if (s.activeCell) attrs += ` activeCell="${escapeXmlAttr(s.activeCell)}"`;
  if (s.sqref) attrs += ` sqref="${escapeXmlAttr(s.sqref)}"`;
  return `<selection${attrs}/>`;
};

const serializeSheetFormatPr = (ws: Worksheet): string => {
  let attrs = '';
  if (ws.baseColWidth !== undefined) attrs += ` baseColWidth="${ws.baseColWidth}"`;
  if (ws.defaultColumnWidth !== undefined) attrs += ` defaultColWidth="${ws.defaultColumnWidth}"`;
  if (ws.defaultRowHeight !== undefined) attrs += ` defaultRowHeight="${ws.defaultRowHeight}"`;
  if (ws.customHeight !== undefined) attrs += ` customHeight="${ws.customHeight ? '1' : '0'}"`;
  if (ws.zeroHeight !== undefined) attrs += ` zeroHeight="${ws.zeroHeight ? '1' : '0'}"`;
  if (ws.thickTop !== undefined) attrs += ` thickTop="${ws.thickTop ? '1' : '0'}"`;
  if (ws.thickBottom !== undefined) attrs += ` thickBottom="${ws.thickBottom ? '1' : '0'}"`;

  // outlineLevelRow / outlineLevelCol — emit explicit if set, else
  // auto-compute the max from row/column dimensions so Excel renders
  // the outline button strip with the right depth.
  const olRow = ws.outlineLevelRow ?? maxOutlineLevel(ws.rowDimensions);
  if (olRow > 0) attrs += ` outlineLevelRow="${olRow}"`;
  const olCol = ws.outlineLevelCol ?? maxOutlineLevel(ws.columnDimensions);
  if (olCol > 0) attrs += ` outlineLevelCol="${olCol}"`;

  if (attrs.length === 0) return '';
  return `<sheetFormatPr${attrs}/>`;
};

const maxOutlineLevel = (
  dims: ReadonlyMap<number, { outlineLevel?: number }>,
): number => {
  let max = 0;
  for (const d of dims.values()) {
    if (d.outlineLevel !== undefined && d.outlineLevel > max) max = d.outlineLevel;
  }
  return max;
};

const serializeCols = (cols: ReadonlyMap<number, ColumnDimension>): string => {
  const sorted = [...cols.values()].sort((a, b) => a.min - b.min);
  const parts: string[] = ['<cols>'];
  for (const dim of sorted) parts.push(serializeColumnDimension(dim));
  parts.push('</cols>');
  return parts.join('');
};

const serializeColumnDimension = (dim: ColumnDimension): string => {
  let attrs = ` min="${dim.min}" max="${dim.max}"`;
  if (dim.width !== undefined) attrs += ` width="${dim.width}"`;
  if (dim.style !== undefined) attrs += ` style="${dim.style}"`;
  if (dim.hidden) attrs += ' hidden="1"';
  if (dim.bestFit) attrs += ' bestFit="1"';
  if (dim.customWidth) attrs += ' customWidth="1"';
  if (dim.outlineLevel !== undefined) attrs += ` outlineLevel="${dim.outlineLevel}"`;
  if (dim.collapsed) attrs += ' collapsed="1"';
  return `<col${attrs}/>`;
};

const serializeAutoFilter = (filter: AutoFilter): string => {
  const refAttr = ` ref="${escapeXmlAttr(filter.ref)}"`;
  if (filter.filterColumns.length === 0) return `<autoFilter${refAttr}/>`;
  const parts: string[] = [`<autoFilter${refAttr}>`];
  for (const fc of filter.filterColumns) {
    parts.push(`<filterColumn colId="${fc.colId}">`);
    let filtersAttrs = '';
    if (fc.blank !== undefined) filtersAttrs += ` blank="${fc.blank ? '1' : '0'}"`;
    if (fc.values.length === 0) {
      parts.push(`<filters${filtersAttrs}/>`);
    } else {
      parts.push(`<filters${filtersAttrs}>`);
      for (const v of fc.values) parts.push(`<filter val="${escapeXmlAttr(v)}"/>`);
      parts.push('</filters>');
    }
    parts.push('</filterColumn>');
  }
  parts.push('</autoFilter>');
  return parts.join('');
};

const serializeConditionalFormatting = (cf: ConditionalFormatting): string => {
  const sqref = escapeXmlAttr(multiCellRangeToString(cf.sqref));
  let attrs = ` sqref="${sqref}"`;
  if (cf.pivot !== undefined) attrs += ` pivot="${cf.pivot ? '1' : '0'}"`;
  if (cf.rules.length === 0) return `<conditionalFormatting${attrs}/>`;
  const parts: string[] = [`<conditionalFormatting${attrs}>`];
  for (const rule of cf.rules) parts.push(serializeCfRule(rule));
  parts.push('</conditionalFormatting>');
  return parts.join('');
};

const serializeCfRule = (rule: ConditionalFormattingRule): string => {
  let attrs = ` type="${rule.type}" priority="${rule.priority}"`;
  if (rule.dxfId !== undefined) attrs += ` dxfId="${rule.dxfId}"`;
  if (rule.stopIfTrue) attrs += ' stopIfTrue="1"';
  if (rule.operator) attrs += ` operator="${escapeXmlAttr(rule.operator)}"`;
  if (rule.text !== undefined) attrs += ` text="${escapeXmlAttr(rule.text)}"`;
  if (rule.percent !== undefined) attrs += ` percent="${rule.percent ? '1' : '0'}"`;
  if (rule.bottom !== undefined) attrs += ` bottom="${rule.bottom ? '1' : '0'}"`;
  if (rule.rank !== undefined) attrs += ` rank="${rule.rank}"`;
  if (rule.aboveAverage !== undefined) attrs += ` aboveAverage="${rule.aboveAverage ? '1' : '0'}"`;
  if (rule.equalAverage !== undefined) attrs += ` equalAverage="${rule.equalAverage ? '1' : '0'}"`;
  if (rule.stdDev !== undefined) attrs += ` stdDev="${rule.stdDev}"`;
  if (rule.timePeriod !== undefined) attrs += ` timePeriod="${rule.timePeriod}"`;

  const inner: string[] = [];
  for (const f of rule.formulas) inner.push(`<formula>${escapeXmlText(f)}</formula>`);
  if (rule.innerXml) inner.push(rule.innerXml);
  if (inner.length === 0) return `<cfRule${attrs}/>`;
  return `<cfRule${attrs}>${inner.join('')}</cfRule>`;
};

const serializeDataValidations = (dvs: ReadonlyArray<DataValidation>): string => {
  const parts: string[] = [`<dataValidations count="${dvs.length}">`];
  for (const dv of dvs) parts.push(serializeDataValidation(dv));
  parts.push('</dataValidations>');
  return parts.join('');
};

const serializeDataValidation = (dv: DataValidation): string => {
  let attrs = ` type="${dv.type}"`;
  if (dv.errorStyle) attrs += ` errorStyle="${dv.errorStyle}"`;
  if (dv.operator) attrs += ` operator="${dv.operator}"`;
  if (dv.allowBlank) attrs += ' allowBlank="1"';
  if (dv.showDropDown) attrs += ' showDropDown="1"';
  if (dv.showInputMessage) attrs += ' showInputMessage="1"';
  if (dv.showErrorMessage) attrs += ' showErrorMessage="1"';
  if (dv.errorTitle !== undefined) attrs += ` errorTitle="${escapeXmlAttr(dv.errorTitle)}"`;
  if (dv.error !== undefined) attrs += ` error="${escapeXmlAttr(dv.error)}"`;
  if (dv.promptTitle !== undefined) attrs += ` promptTitle="${escapeXmlAttr(dv.promptTitle)}"`;
  if (dv.prompt !== undefined) attrs += ` prompt="${escapeXmlAttr(dv.prompt)}"`;
  attrs += ` sqref="${escapeXmlAttr(multiCellRangeToString(dv.sqref))}"`;

  const formulas: string[] = [];
  if (dv.formula1 !== undefined) formulas.push(`<formula1>${escapeXmlText(dv.formula1)}</formula1>`);
  if (dv.formula2 !== undefined) formulas.push(`<formula2>${escapeXmlText(dv.formula2)}</formula2>`);
  if (formulas.length === 0) return `<dataValidation${attrs}/>`;
  return `<dataValidation${attrs}>${formulas.join('')}</dataValidation>`;
};

const serializeSheetProtection = (sp: SheetProtection): string | undefined => {
  const flagKeys = [
    'sheet',
    'objects',
    'scenarios',
    'formatCells',
    'formatColumns',
    'formatRows',
    'insertColumns',
    'insertRows',
    'insertHyperlinks',
    'deleteColumns',
    'deleteRows',
    'selectLockedCells',
    'selectUnlockedCells',
    'sort',
    'autoFilter',
    'pivotTables',
  ] as const;
  let attrs = '';
  for (const k of flagKeys) {
    const v = sp[k];
    if (v !== undefined) attrs += ` ${k}="${v ? '1' : '0'}"`;
  }
  if (sp.algorithmName !== undefined) attrs += ` algorithmName="${escapeXmlAttr(sp.algorithmName)}"`;
  if (sp.hashValue !== undefined) attrs += ` hashValue="${escapeXmlAttr(sp.hashValue)}"`;
  if (sp.saltValue !== undefined) attrs += ` saltValue="${escapeXmlAttr(sp.saltValue)}"`;
  if (sp.spinCount !== undefined) attrs += ` spinCount="${sp.spinCount}"`;
  if (attrs.length === 0) return undefined;
  return `<sheetProtection${attrs}/>`;
};

const serializeSheetProperties = (sp: SheetProperties): string | undefined => {
  let attrs = '';
  if (sp.codeName !== undefined) attrs += ` codeName="${escapeXmlAttr(sp.codeName)}"`;
  if (sp.enableFormatConditionsCalculation === false) attrs += ' enableFormatConditionsCalculation="0"';
  if (sp.enableFormatConditionsCalculation === true) attrs += ' enableFormatConditionsCalculation="1"';
  if (sp.filterMode !== undefined) attrs += ` filterMode="${sp.filterMode ? '1' : '0'}"`;
  if (sp.published !== undefined) attrs += ` published="${sp.published ? '1' : '0'}"`;
  if (sp.syncHorizontal !== undefined) attrs += ` syncHorizontal="${sp.syncHorizontal ? '1' : '0'}"`;
  if (sp.syncRef !== undefined) attrs += ` syncRef="${escapeXmlAttr(sp.syncRef)}"`;
  if (sp.syncVertical !== undefined) attrs += ` syncVertical="${sp.syncVertical ? '1' : '0'}"`;
  if (sp.transitionEvaluation !== undefined) attrs += ` transitionEvaluation="${sp.transitionEvaluation ? '1' : '0'}"`;
  if (sp.transitionEntry !== undefined) attrs += ` transitionEntry="${sp.transitionEntry ? '1' : '0'}"`;

  const children: string[] = [];
  if (sp.tabColor) {
    let tcAttrs = '';
    if (sp.tabColor.rgb !== undefined) tcAttrs += ` rgb="${escapeXmlAttr(sp.tabColor.rgb)}"`;
    if (sp.tabColor.indexed !== undefined) tcAttrs += ` indexed="${sp.tabColor.indexed}"`;
    if (sp.tabColor.theme !== undefined) tcAttrs += ` theme="${sp.tabColor.theme}"`;
    if (sp.tabColor.auto !== undefined) tcAttrs += ` auto="${sp.tabColor.auto ? '1' : '0'}"`;
    if (sp.tabColor.tint !== undefined) tcAttrs += ` tint="${sp.tabColor.tint}"`;
    children.push(`<tabColor${tcAttrs}/>`);
  }
  if (sp.outlinePr) {
    let opAttrs = '';
    if (sp.outlinePr.applyStyles !== undefined) opAttrs += ` applyStyles="${sp.outlinePr.applyStyles ? '1' : '0'}"`;
    if (sp.outlinePr.summaryBelow !== undefined) opAttrs += ` summaryBelow="${sp.outlinePr.summaryBelow ? '1' : '0'}"`;
    if (sp.outlinePr.summaryRight !== undefined) opAttrs += ` summaryRight="${sp.outlinePr.summaryRight ? '1' : '0'}"`;
    if (sp.outlinePr.showOutlineSymbols !== undefined)
      opAttrs += ` showOutlineSymbols="${sp.outlinePr.showOutlineSymbols ? '1' : '0'}"`;
    children.push(`<outlinePr${opAttrs}/>`);
  }
  if (sp.pageSetUpPr) {
    let psAttrs = '';
    if (sp.pageSetUpPr.autoPageBreaks !== undefined)
      psAttrs += ` autoPageBreaks="${sp.pageSetUpPr.autoPageBreaks ? '1' : '0'}"`;
    if (sp.pageSetUpPr.fitToPage !== undefined) psAttrs += ` fitToPage="${sp.pageSetUpPr.fitToPage ? '1' : '0'}"`;
    children.push(`<pageSetUpPr${psAttrs}/>`);
  }

  if (attrs.length === 0 && children.length === 0) return undefined;
  if (children.length === 0) return `<sheetPr${attrs}/>`;
  return `<sheetPr${attrs}>${children.join('')}</sheetPr>`;
};

const serializeCellWatches = (watches: ReadonlyArray<CellWatch>): string => {
  const parts: string[] = ['<cellWatches>'];
  for (const w of watches) parts.push(`<cellWatch r="${escapeXmlAttr(w.ref)}"/>`);
  parts.push('</cellWatches>');
  return parts.join('');
};

const serializeIgnoredErrors = (errs: ReadonlyArray<IgnoredError>): string => {
  const parts: string[] = ['<ignoredErrors>'];
  for (const ie of errs) {
    let attrs = ` sqref="${escapeXmlAttr(multiCellRangeToString(ie.sqref))}"`;
    if (ie.evalError) attrs += ' evalError="1"';
    if (ie.twoDigitTextYear) attrs += ' twoDigitTextYear="1"';
    if (ie.numberStoredAsText) attrs += ' numberStoredAsText="1"';
    if (ie.formula) attrs += ' formula="1"';
    if (ie.formulaRange) attrs += ' formulaRange="1"';
    if (ie.unlockedFormula) attrs += ' unlockedFormula="1"';
    if (ie.emptyCellReference) attrs += ' emptyCellReference="1"';
    if (ie.listDataValidation) attrs += ' listDataValidation="1"';
    if (ie.calculatedColumn) attrs += ' calculatedColumn="1"';
    parts.push(`<ignoredError${attrs}/>`);
  }
  parts.push('</ignoredErrors>');
  return parts.join('');
};

const serializeHyperlinks = (links: ReadonlyArray<Hyperlink>, rels: Relationships | undefined): string => {
  const parts: string[] = ['<hyperlinks>'];
  for (const link of links) {
    let attrs = ` ref="${escapeXmlAttr(link.ref)}"`;
    if (link.target !== undefined) {
      // Allocate or reuse a rels entry. We need rels to host external URLs;
      // when ctx.rels is missing we fall back to inlining via location only.
      if (rels) {
        let rId = link.rId;
        if (!rId) {
          // Find the next free rIdN that doesn't already exist on this
          // sheet's rels. Pre-loaded relsExtras can occupy low-numbered
          // ids; a naive `rId${len+1}` would collide.
          let n = rels.rels.length + 1;
          rId = `rId${n}`;
          while (rels.rels.some((r) => r.id === rId)) {
            n++;
            rId = `rId${n}`;
          }
        }
        // Add rel only if no entry already targets this URL (conservative).
        if (!rels.rels.some((r) => r.id === rId)) {
          rels.rels.push({
            id: rId,
            type: HYPERLINK_REL_TYPE,
            target: link.target,
            targetMode: 'External',
          });
        }
        attrs += ` r:id="${escapeXmlAttr(rId)}"`;
      }
    }
    if (link.location !== undefined) attrs += ` location="${escapeXmlAttr(link.location)}"`;
    if (link.tooltip !== undefined) attrs += ` tooltip="${escapeXmlAttr(link.tooltip)}"`;
    if (link.display !== undefined) attrs += ` display="${escapeXmlAttr(link.display)}"`;
    parts.push(`<hyperlink${attrs}/>`);
  }
  parts.push('</hyperlinks>');
  return parts.join('');
};

const serializeRowDimensionAttrs = (dim: RowDimension): string => {
  let attrs = '';
  if (dim.height !== undefined) attrs += ` ht="${dim.height}"`;
  if (dim.customHeight) attrs += ' customHeight="1"';
  if (dim.hidden) attrs += ' hidden="1"';
  if (dim.outlineLevel !== undefined) attrs += ` outlineLevel="${dim.outlineLevel}"`;
  if (dim.collapsed) attrs += ' collapsed="1"';
  if (dim.style !== undefined) attrs += ` s="${dim.style}" customFormat="1"`;
  return attrs;
};

/**
 * Serialise a captured worksheet body extra (XmlNode) for inline emit.
 * Reuses the namespace-aware serializer then strips the XML declaration.
 * Excel tolerates the redundant per-element `xmlns="…"` declarations
 * that arise from emitting each child as its own root.
 */
const serializeBodyExtraNode = (node: XmlNode): string =>
  new TextDecoder().decode(serializeXml(node, { xmlDeclaration: false }));
