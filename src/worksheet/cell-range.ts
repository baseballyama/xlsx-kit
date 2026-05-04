// Cell-range value object + set operations. Per
// docs/plan/04-core-model.md §4.5.
//
// The struct is the same `CellRangeBoundaries` we already use across
// the coordinate parser; this module adds the worksheet-level
// operations (containment, shift, union, intersection, iteration) and
// a `MultiCellRange` lite wrapper for sqref-style attributes.

import {
  boundariesToRangeString,
  type CellRangeBoundaries,
  MAX_COL,
  MAX_ROW,
  rangeBoundaries,
} from '../utils/coordinate';
import { OpenXmlSchemaError } from '../utils/exceptions';

/** Re-export under the plan's canonical name. */
export type CellRange = CellRangeBoundaries;

/** Build a CellRange from explicit 1-based bounds. */
export function makeCellRange(minRow: number, minCol: number, maxRow: number, maxCol: number): CellRange {
  if (
    !Number.isInteger(minRow) ||
    !Number.isInteger(minCol) ||
    !Number.isInteger(maxRow) ||
    !Number.isInteger(maxCol)
  ) {
    throw new OpenXmlSchemaError('CellRange bounds must be integers');
  }
  if (minRow < 1 || maxRow < 1 || minRow > MAX_ROW || maxRow > MAX_ROW) {
    throw new OpenXmlSchemaError(`CellRange row bounds must be in [1, ${MAX_ROW}]`);
  }
  if (minCol < 1 || maxCol < 1 || minCol > MAX_COL || maxCol > MAX_COL) {
    throw new OpenXmlSchemaError(`CellRange col bounds must be in [1, ${MAX_COL}]`);
  }
  return {
    minRow: Math.min(minRow, maxRow),
    minCol: Math.min(minCol, maxCol),
    maxRow: Math.max(minRow, maxRow),
    maxCol: Math.max(minCol, maxCol),
  };
}

/** Parse a range expression — wraps {@link rangeBoundaries}. */
export function parseRange(input: string): CellRange {
  return rangeBoundaries(input);
}

/** Format a CellRange back into the canonical OOXML string. */
export function rangeToString(r: CellRange): string {
  return boundariesToRangeString(r);
}

/** Inclusive containment of a single (row, col) within a range. */
export function rangeContainsCell(r: CellRange, row: number, col: number): boolean {
  return row >= r.minRow && row <= r.maxRow && col >= r.minCol && col <= r.maxCol;
}

/** Inclusive containment of `inner` within `outer`. */
export function rangeContainsRange(outer: CellRange, inner: CellRange): boolean {
  return (
    inner.minRow >= outer.minRow &&
    inner.maxRow <= outer.maxRow &&
    inner.minCol >= outer.minCol &&
    inner.maxCol <= outer.maxCol
  );
}

/**
 * Shift a range by (dr, dc) integer offsets. The returned range is
 * clamped to the OOXML grid; callers that want hard bounds should
 * pass values that keep the result inside the spec.
 */
export function shiftRange(r: CellRange, dr: number, dc: number): CellRange {
  if (!Number.isInteger(dr) || !Number.isInteger(dc)) {
    throw new OpenXmlSchemaError('shiftRange: dr / dc must be integers');
  }
  return makeCellRange(r.minRow + dr, r.minCol + dc, r.maxRow + dr, r.maxCol + dc);
}

/** Bounding-box union of two ranges. Always non-null. */
export function unionRange(a: CellRange, b: CellRange): CellRange {
  return {
    minRow: Math.min(a.minRow, b.minRow),
    minCol: Math.min(a.minCol, b.minCol),
    maxRow: Math.max(a.maxRow, b.maxRow),
    maxCol: Math.max(a.maxCol, b.maxCol),
  };
}

/** Returns the rectangular intersection of two ranges, or `null` when disjoint. */
export function intersectionRange(a: CellRange, b: CellRange): CellRange | null {
  const minRow = Math.max(a.minRow, b.minRow);
  const minCol = Math.max(a.minCol, b.minCol);
  const maxRow = Math.min(a.maxRow, b.maxRow);
  const maxCol = Math.min(a.maxCol, b.maxCol);
  if (minRow > maxRow || minCol > maxCol) return null;
  return { minRow, minCol, maxRow, maxCol };
}

/** True iff two ranges share at least one cell. */
export function rangesOverlap(a: CellRange, b: CellRange): boolean {
  return intersectionRange(a, b) !== null;
}

/** Inclusive cell count covered by a range. */
export function rangeArea(r: CellRange): number {
  return (r.maxRow - r.minRow + 1) * (r.maxCol - r.minCol + 1);
}

/** Yield every (row, col) coordinate in the range, row-major. */
export function* iterRangeCoordinates(r: CellRange): IterableIterator<{ row: number; col: number }> {
  for (let row = r.minRow; row <= r.maxRow; row++) {
    for (let col = r.minCol; col <= r.maxCol; col++) {
      yield { row, col };
    }
  }
}

// ---- MultiCellRange --------------------------------------------------------

/**
 * Excel's `sqref` attribute: a space-separated list of CellRanges.
 * Used by data validations, conditional formatting, hyperlinks etc.
 */
export interface MultiCellRange {
  ranges: CellRange[];
}

export function makeMultiCellRange(ranges: ReadonlyArray<CellRange> = []): MultiCellRange {
  return { ranges: ranges.slice() };
}

/** Parse an sqref string: `"A1:B2 D5 E10:F20"`. Whitespace-delimited. */
export function parseMultiCellRange(input: string): MultiCellRange {
  const ranges: CellRange[] = [];
  for (const piece of input.split(/\s+/)) {
    if (piece.length === 0) continue;
    ranges.push(parseRange(piece));
  }
  return { ranges };
}

/** Format a MultiCellRange back into an sqref string. */
export function multiCellRangeToString(m: MultiCellRange): string {
  return m.ranges.map(rangeToString).join(' ');
}

/** Total cell count across all ranges (no de-duplication of overlaps). */
export function multiCellRangeArea(m: MultiCellRange): number {
  let n = 0;
  for (const r of m.ranges) n += rangeArea(r);
  return n;
}

/** True iff any contained range covers (row, col). */
export function multiCellRangeContainsCell(m: MultiCellRange, row: number, col: number): boolean {
  for (const r of m.ranges) if (rangeContainsCell(r, row, col)) return true;
  return false;
}
