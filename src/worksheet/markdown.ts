// Worksheet → GitHub-Flavored Markdown table renderer. Sibling of
// worksheetToHtml for the markdown output side.
//
// Markdown tables are flat: no rowspan/colspan, no styles. Merged
// cells contribute their value at the top-left only (other slots
// render as empty cells). Cell values are pipe-escaped (`\|`) and
// newlines are replaced with `<br>` to keep one-row-per-line.

import { cellValueAsString } from '../cell/cell';
import { boundariesToRangeString } from '../utils/coordinate';
import { parseRange } from './cell-range';
import { getCell, getDataExtent, getMergedCells, type Worksheet } from './worksheet';

const escapeMarkdownCell = (s: string): string =>
  s
    .replace(/\\/g, '\\\\')
    .replace(/\|/g, '\\|')
    .replace(/\r?\n/g, '<br>');

/**
 * Render a worksheet range as a GitHub-Flavored Markdown table. The
 * first row of the range becomes the header. Output is one row per
 * line, with a header-separator line (`| --- | --- |`) below the
 * header.
 *
 * Merged ranges are flattened: the top-left cell of each merge keeps
 * its value, other slots render as empty cells (markdown has no
 * rowspan / colspan).
 *
 * Cell values are pipe-escaped (`\|`) and newlines become `<br>` so
 * the table stays on one row per source spreadsheet row.
 *
 * Returns `''` when the range covers zero rows or columns.
 */
export function worksheetToMarkdownTable(ws: Worksheet, range: string): string {
  const { minRow, minCol, maxRow, maxCol } = parseRange(range);
  const colCount = maxCol - minCol + 1;
  if (colCount <= 0 || maxRow < minRow) return '';
  // Build the skip set for merged-cell non-top-left slots.
  const skip = new Set<string>();
  for (const m of getMergedCells(ws)) {
    if (m.maxRow < minRow || m.minRow > maxRow || m.maxCol < minCol || m.minCol > maxCol) continue;
    const tlRow = Math.max(m.minRow, minRow);
    const tlCol = Math.max(m.minCol, minCol);
    const brRow = Math.min(m.maxRow, maxRow);
    const brCol = Math.min(m.maxCol, maxCol);
    for (let r = tlRow; r <= brRow; r++) {
      for (let c = tlCol; c <= brCol; c++) {
        if (r === tlRow && c === tlCol) continue;
        skip.add(`${r},${c}`);
      }
    }
  }
  const renderRow = (r: number): string => {
    const cells: string[] = [];
    for (let c = minCol; c <= maxCol; c++) {
      if (skip.has(`${r},${c}`)) {
        cells.push('');
        continue;
      }
      const cell = getCell(ws, r, c);
      cells.push(cell ? escapeMarkdownCell(cellValueAsString(cell.value)) : '');
    }
    return `| ${cells.join(' | ')} |`;
  };
  const lines: string[] = [];
  // Header row.
  lines.push(renderRow(minRow));
  // Header separator (one '---' per column).
  lines.push(`| ${new Array(colCount).fill('---').join(' | ')} |`);
  // Data rows.
  for (let r = minRow + 1; r <= maxRow; r++) {
    lines.push(renderRow(r));
  }
  return lines.join('\n');
}

/**
 * Whole-worksheet shortcut over {@link worksheetToMarkdownTable}:
 * serialises the sheet's data extent (`getDataExtent`) as a GFM
 * markdown table. Returns `''` for an empty worksheet (mirrors the
 * CSV / HTML shortcut conventions).
 */
export function getWorksheetAsMarkdownTable(ws: Worksheet): string {
  const ext = getDataExtent(ws);
  if (!ext) return '';
  return worksheetToMarkdownTable(ws, boundariesToRangeString(ext));
}
