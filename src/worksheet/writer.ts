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
import { escapeCellString } from '../utils/escape';
import { OpenXmlSchemaError } from '../utils/exceptions';
import type { SharedStringsTable } from '../workbook/shared-strings';
import { addSharedString } from '../workbook/shared-strings';
import { SHEET_MAIN_NS } from '../xml/namespaces';
import { rangeToString } from './cell-range';
import type { Worksheet } from './worksheet';

export interface WorksheetWriteContext {
  /** Accumulator the writer mutates as it emits string cells. */
  sharedStrings: SharedStringsTable;
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
    serializeDimension(ws),
    '<sheetData>',
  ];
  // Iterate rows in numeric order so writer output is deterministic.
  const rowKeys = [...ws.rows.keys()].sort((a, b) => a - b);
  for (const rowIdx of rowKeys) {
    const row = ws.rows.get(rowIdx);
    if (!row || row.size === 0) continue;
    const colKeys = [...row.keys()].sort((a, b) => a - b);
    parts.push(`<row r="${rowIdx}">`);
    for (const colIdx of colKeys) {
      const cell = row.get(colIdx);
      if (cell) parts.push(serializeCell(cell, ctx));
    }
    parts.push('</row>');
  }
  parts.push('</sheetData>');
  if (ws.mergedCells.length > 0) {
    parts.push(`<mergeCells count="${ws.mergedCells.length}">`);
    for (const range of ws.mergedCells) {
      parts.push(`<mergeCell ref="${rangeToString(range)}"/>`);
    }
    parts.push('</mergeCells>');
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

const serializeCell = (cell: Cell, ctx: WorksheetWriteContext): string => {
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
    // Excel stores dates as serial numbers. The conversion is in utils/datetime.
    // For stage-1 we throw — date round-trip needs the styleId-driven format
    // detection that lives in §5.5 (deferred).
    throw new OpenXmlSchemaError(
      `worksheet: Date cells require the date-format integration (deferred to §5.5); cell ${ref}`,
    );
  }
  if (typeof value === 'object' && value !== null && (value as { kind?: string }).kind === 'duration') {
    // Same comment as Date — deferred.
    throw new OpenXmlSchemaError(
      `worksheet: duration cells require the duration-format integration (deferred); cell ${ref}`,
    );
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
