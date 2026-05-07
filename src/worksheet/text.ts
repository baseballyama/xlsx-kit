// Worksheet → ASCII-art text table renderer. Sibling of
// worksheetToHtml / worksheetToMarkdownTable for plain-text output.
//
// Each column is padded to its widest cell width; cells are joined
// with ` | ` and each row wrapped in `| ... |`. A `+---+---+` border
// separates header from data rows. Merged cells are flattened (top-
// left keeps its value, others are empty).

import { cellValueAsString } from '../cell/cell';
import { boundariesToRangeString } from '../utils/coordinate';
import { parseRange } from './cell-range';
import { getCell, getDataExtent, getMergedCells, type Worksheet } from './worksheet';

/**
 * Render a worksheet range as a plain ASCII-art table. The first
 * row of the range is treated as the header and separated from
 * subsequent rows by a `+---+---+` border. Each column is padded
 * (with spaces) to the width of its widest cell so values line up
 * vertically.
 *
 * Merged ranges are flattened — the top-left cell of each merge
 * keeps its value, others render as empty cells (text tables have
 * no rowspan / colspan equivalent).
 *
 * Newlines inside cell values are replaced with a space so the
 * one-row-per-line invariant holds.
 *
 * Returns `''` when the range covers zero rows or columns.
 */
export function worksheetToTextTable(ws: Worksheet, range: string): string {
  const { minRow, minCol, maxRow, maxCol } = parseRange(range);
  const colCount = maxCol - minCol + 1;
  if (colCount <= 0 || maxRow < minRow) return '';
  // Build the skip set for merge non-top-left slots.
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
  // Materialise the grid as strings with newlines flattened.
  const grid: string[][] = [];
  for (let r = minRow; r <= maxRow; r++) {
    const row: string[] = [];
    for (let c = minCol; c <= maxCol; c++) {
      if (skip.has(`${r},${c}`)) {
        row.push('');
        continue;
      }
      const cell = getCell(ws, r, c);
      const v = cell ? cellValueAsString(cell.value) : '';
      row.push(v.replace(/\r?\n/g, ' '));
    }
    grid.push(row);
  }
  // Determine per-column widths.
  const widths = new Array<number>(colCount).fill(0);
  for (const row of grid) {
    for (let j = 0; j < colCount; j++) {
      const cell = row[j] ?? '';
      const w = widths[j] ?? 0;
      if (cell.length > w) widths[j] = cell.length;
    }
  }
  const renderRow = (row: string[]): string => {
    const parts: string[] = [];
    for (let j = 0; j < colCount; j++) {
      parts.push((row[j] ?? '').padEnd(widths[j] ?? 0, ' '));
    }
    return `| ${parts.join(' | ')} |`;
  };
  const sep = `+${widths.map((w) => '-'.repeat(w + 2)).join('+')}+`;
  const lines: string[] = [];
  // Header row.
  const firstRow = grid[0];
  if (firstRow) lines.push(renderRow(firstRow));
  lines.push(sep);
  for (let i = 1; i < grid.length; i++) {
    const row = grid[i];
    if (row) lines.push(renderRow(row));
  }
  return lines.join('\n');
}

/**
 * Whole-worksheet shortcut over {@link worksheetToTextTable}: serialises
 * the sheet's data extent (`getDataExtent`) as an ASCII-art table.
 * Returns `''` for an empty worksheet (mirrors the CSV / HTML /
 * Markdown shortcut conventions).
 */
export function getWorksheetAsTextTable(ws: Worksheet): string {
  const ext = getDataExtent(ws);
  if (!ext) return '';
  return worksheetToTextTable(ws, boundariesToRangeString(ext));
}
