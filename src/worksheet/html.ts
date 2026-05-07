// Worksheet → HTML rendering. The natural payoff for the *ToCss
// helper family (color/font/fill/border/alignment/cellStyle) +
// cssRecordToInlineStyle.
//
// Intentionally narrow: produces a plain `<table>` with `<tr>` /
// `<td>` and inline `style="…"` attributes per cell, with merged
// ranges collapsed via `rowspan` / `colspan`. No CSS classes, no
// `<thead>` / `<tbody>` split — callers can wrap / re-style as
// needed.

import { cellStyleToCss } from '../styles/cell-style';
import { boundariesToRangeString } from '../utils/coordinate';
import { cssRecordToInlineStyle } from '../utils/css';
import { cellValueAsString } from '../cell/cell';
import type { Workbook } from '../workbook/workbook';
import { parseRange } from './cell-range';
import { getCell, getDataExtent, getMergedCells, type Worksheet } from './worksheet';

const escapeHtml = (s: string): string =>
  s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

/**
 * Render a worksheet range as an HTML `<table>`. Each cell becomes
 * a `<td>` with an inline `style="…"` attribute computed from the
 * cell's resolved {@link cellStyleToCss}. Merged ranges that
 * overlap the target range collapse to the top-left cell with
 * `rowspan` / `colspan`; the other slots are skipped (no `<td>`).
 *
 * Empty cells produce an empty `<td>` (still styled). Cell values
 * are HTML-escaped.
 */
export function worksheetToHtml(wb: Workbook, ws: Worksheet, range: string): string {
  const { minRow, minCol, maxRow, maxCol } = parseRange(range);
  // Build a "skip" set: every (row, col) that's covered by a merge
  // but isn't the merge top-left.
  const skip = new Set<string>();
  // And a "topLeft" map: (row, col) → { rowspan, colspan }.
  const merges = new Map<string, { rowspan: number; colspan: number }>();
  for (const m of getMergedCells(ws)) {
    if (m.maxRow < minRow || m.minRow > maxRow || m.maxCol < minCol || m.minCol > maxCol) continue;
    const tlRow = Math.max(m.minRow, minRow);
    const tlCol = Math.max(m.minCol, minCol);
    const brRow = Math.min(m.maxRow, maxRow);
    const brCol = Math.min(m.maxCol, maxCol);
    merges.set(`${tlRow},${tlCol}`, {
      rowspan: brRow - tlRow + 1,
      colspan: brCol - tlCol + 1,
    });
    for (let r = tlRow; r <= brRow; r++) {
      for (let c = tlCol; c <= brCol; c++) {
        if (r === tlRow && c === tlCol) continue;
        skip.add(`${r},${c}`);
      }
    }
  }

  const lines: string[] = ['<table>'];
  for (let r = minRow; r <= maxRow; r++) {
    lines.push('<tr>');
    for (let c = minCol; c <= maxCol; c++) {
      const key = `${r},${c}`;
      if (skip.has(key)) continue;
      const cell = getCell(ws, r, c);
      const styleObj = cell ? cellStyleToCss(wb, cell) : {};
      const style = cssRecordToInlineStyle(styleObj);
      const merge = merges.get(key);
      const attrs: string[] = [];
      if (style.length > 0) attrs.push(`style="${escapeHtml(style)}"`);
      if (merge && merge.rowspan > 1) attrs.push(`rowspan="${merge.rowspan}"`);
      if (merge && merge.colspan > 1) attrs.push(`colspan="${merge.colspan}"`);
      const open = attrs.length > 0 ? `<td ${attrs.join(' ')}>` : '<td>';
      const text = cell ? escapeHtml(cellValueAsString(cell.value)) : '';
      lines.push(`${open}${text}</td>`);
    }
    lines.push('</tr>');
  }
  lines.push('</table>');
  return lines.join('');
}

/**
 * Whole-worksheet shortcut over {@link worksheetToHtml}: serialises
 * the sheet's data extent (`getDataExtent`) as an HTML `<table>`.
 * Returns `''` for an empty worksheet (mirrors
 * {@link getWorksheetAsCsv}'s convention).
 */
export function getWorksheetAsHtml(wb: Workbook, ws: Worksheet): string {
  const ext = getDataExtent(ws);
  if (!ext) return '';
  return worksheetToHtml(wb, ws, boundariesToRangeString(ext));
}
