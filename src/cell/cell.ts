// Cell value model. Mirrors openpyxl/openpyxl/cell/cell.py.
//
// A Cell is a plain mutable object: the worksheet stores millions of
// these, so per-cell freezes / spreads aren't viable on the hot path.
// Per docs/plan/04-core-model.md §2 the public surface stays small —
// makeCell + getCoordinate + targeted setters — and uses
// discriminated unions for the special CellValue shapes (formula,
// rich text, duration, error).

import { columnLetterFromIndex, MAX_COL, MAX_ROW } from '../utils/coordinate';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { ERROR_CODES } from '../utils/inference';
import type { RichText } from './rich-text';

/** Excel error tokens. */
export type ExcelErrorCode = '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#GETTING_DATA';

/** Formula sub-kind — drives the OOXML `<f t="…">` attribute. */
export type FormulaKind = 'normal' | 'array' | 'shared' | 'dataTable';

export interface FormulaValue {
  readonly kind: 'formula';
  readonly formula: string;
  readonly t: FormulaKind;
  /** Cached value Excel last computed for the cell, used when `data_only` reads it back. */
  readonly cachedValue?: number | string | boolean;
  /** Range string (`"A1:A10"`) for array / shared / dataTable formulas. */
  readonly ref?: string;
  /** Shared-formula index. */
  readonly si?: number;
  /** Data-table specific fields (mirrors openpyxl DataTableFormula). */
  readonly r1?: string;
  readonly r2?: string;
  readonly dt2D?: boolean;
  readonly dtr?: boolean;
  readonly del1?: boolean;
  readonly del2?: boolean;
  readonly aca?: boolean;
  readonly ca?: boolean;
}

export type CellValue =
  | number
  | string
  | boolean
  | Date
  | { kind: 'duration'; ms: number }
  | { kind: 'error'; code: ExcelErrorCode }
  | { kind: 'rich-text'; runs: RichText }
  | FormulaValue
  | null;

export interface Cell {
  /** 1-based row index. */
  row: number;
  /** 1-based column index. */
  col: number;
  /** Effective cell value. `null` represents an empty cell. */
  value: CellValue;
  /** Index into Workbook.styles.cellXfs. 0 = default. */
  styleId: number;
  /** Optional reference to a Hyperlink registered on the worksheet. */
  hyperlinkId?: number;
  /** Optional reference to a Comment registered on the worksheet. */
  commentId?: number;
}

/** Marker subtype for the placeholder cells inside a merged range (top-left holds the value). */
export interface MergedCell extends Cell {
  merged: true;
}

const validateCoord = (row: number, col: number): void => {
  if (!Number.isInteger(row) || row < 1 || row > MAX_ROW) {
    throw new OpenXmlSchemaError(`Cell row ${row} out of range [1, ${MAX_ROW}]`);
  }
  if (!Number.isInteger(col) || col < 1 || col > MAX_COL) {
    throw new OpenXmlSchemaError(`Cell col ${col} out of range [1, ${MAX_COL}]`);
  }
};

/** Build a fresh Cell. Validates coordinates against the OOXML grid bounds. */
export function makeCell(row: number, col: number, value: CellValue = null, styleId = 0): Cell {
  validateCoord(row, col);
  return { row, col, value, styleId };
}

/** Format a Cell's coordinate as the canonical "A1" string. */
export function getCoordinate(c: Cell): string {
  return `${columnLetterFromIndex(c.col)}${c.row}`;
}

/**
 * Direct value setter. No type inference, no validation beyond the
 * union — the caller is in charge. Use {@link bindValue} for the
 * "do what I mean" path.
 */
export function setCellValue(c: Cell, value: CellValue): void {
  c.value = value;
}

/**
 * "Smart" setter: infers the cell value from a JS runtime value.
 * - `string` starting with `=` → formula
 * - `string` matching an Excel error token → error variant
 * - other primitives / Date / null pass through verbatim
 *
 * Intentionally not the default — explicit is clearer for typed code,
 * and inferring on every write costs measurable time on the worksheet
 * write hot path.
 */
export function bindValue(c: Cell, value: number | string | boolean | Date | null): void {
  if (typeof value === 'string') {
    if (value.length > 0 && value.charCodeAt(0) === 61 /* '=' */) {
      setFormula(c, value.slice(1));
      return;
    }
    if (ERROR_CODES.has(value)) {
      c.value = { kind: 'error', code: value as ExcelErrorCode };
      return;
    }
    c.value = value;
    return;
  }
  c.value = value;
}

// ---- formula setters -------------------------------------------------------

/** Plain `=A1+B1` style formula. Cached value is optional but recommended for round-trip. */
export function setFormula(c: Cell, formula: string, opts?: { cachedValue?: FormulaValue['cachedValue'] }): void {
  const v: FormulaValue = {
    kind: 'formula',
    t: 'normal',
    formula,
    ...(opts?.cachedValue !== undefined ? { cachedValue: opts.cachedValue } : {}),
  };
  c.value = v;
}

/** Array (CSE) formula spanning a `ref` range. */
export function setArrayFormula(
  c: Cell,
  ref: string,
  formula: string,
  opts?: { cachedValue?: FormulaValue['cachedValue'] },
): void {
  const v: FormulaValue = {
    kind: 'formula',
    t: 'array',
    formula,
    ref,
    ...(opts?.cachedValue !== undefined ? { cachedValue: opts.cachedValue } : {}),
  };
  c.value = v;
}

/**
 * Shared formula. The first cell in the group carries the formula text
 * + ref; subsequent cells with the same `si` carry only the index and
 * Excel reconstructs the formula via reference shifting.
 */
export function setSharedFormula(
  c: Cell,
  si: number,
  formula?: string,
  ref?: string,
  opts?: { cachedValue?: FormulaValue['cachedValue'] },
): void {
  if (!Number.isInteger(si) || si < 0) {
    throw new OpenXmlSchemaError(`setSharedFormula: si must be a non-negative integer; got ${si}`);
  }
  const v: FormulaValue = {
    kind: 'formula',
    t: 'shared',
    formula: formula ?? '',
    si,
    ...(ref !== undefined ? { ref } : {}),
    ...(opts?.cachedValue !== undefined ? { cachedValue: opts.cachedValue } : {}),
  };
  c.value = v;
}

/**
 * Excel data-table formula (`<f t="dataTable">`). These appear as the
 * "What-if Analysis → Data Table" feature output: a 1- or 2-variable
 * sensitivity grid where the formula references one or two input
 * cells. The wire format mirrors openpyxl's `DataTableFormula`:
 *
 * - `ref`     — inclusive cell range the formula spans.
 * - `r1`, `r2`— input cell coordinates ("$A$1" etc.).
 * - `dt2D`    — true for two-variable tables (uses both r1 and r2).
 * - `dtr`     — row-direction flag (true) vs column-direction (false).
 * - `del1`/`del2` — Excel marks one of these true when the
 *   corresponding input cell has been deleted; the formula keeps
 *   round-tripping so Excel can show the warning state.
 * - `aca`/`ca` — alwaysCalculate / calculate flags.
 */
export interface DataTableFormulaOpts {
  ref: string;
  r1?: string;
  r2?: string;
  dt2D?: boolean;
  dtr?: boolean;
  del1?: boolean;
  del2?: boolean;
  aca?: boolean;
  ca?: boolean;
  cachedValue?: FormulaValue['cachedValue'];
}

/**
 * Set a data-table formula on a cell. Preserves all the dt-specific
 * attributes so the writer can re-emit `<f t="dataTable" r1="..." />`
 * verbatim and Excel keeps treating the cell as a Data Table cell.
 */
export function setDataTableFormula(c: Cell, formula: string, opts: DataTableFormulaOpts): void {
  const v: FormulaValue = {
    kind: 'formula',
    t: 'dataTable',
    formula,
    ref: opts.ref,
    ...(opts.r1 !== undefined ? { r1: opts.r1 } : {}),
    ...(opts.r2 !== undefined ? { r2: opts.r2 } : {}),
    ...(opts.dt2D !== undefined ? { dt2D: opts.dt2D } : {}),
    ...(opts.dtr !== undefined ? { dtr: opts.dtr } : {}),
    ...(opts.del1 !== undefined ? { del1: opts.del1 } : {}),
    ...(opts.del2 !== undefined ? { del2: opts.del2 } : {}),
    ...(opts.aca !== undefined ? { aca: opts.aca } : {}),
    ...(opts.ca !== undefined ? { ca: opts.ca } : {}),
    ...(opts.cachedValue !== undefined ? { cachedValue: opts.cachedValue } : {}),
  };
  c.value = v;
}

// ---- value-shape helpers ---------------------------------------------------

/** Build a `{ kind: 'error', code }` cell value. */
export function makeErrorValue(code: ExcelErrorCode): { kind: 'error'; code: ExcelErrorCode } {
  if (!ERROR_CODES.has(code)) {
    throw new OpenXmlSchemaError(`makeErrorValue: unknown error code "${code}"`);
  }
  return Object.freeze({ kind: 'error', code });
}

/** Build a `{ kind: 'duration', ms }` cell value. */
export function makeDurationValue(ms: number): { kind: 'duration'; ms: number } {
  if (!Number.isFinite(ms)) {
    throw new OpenXmlSchemaError(`makeDurationValue: ms "${ms}" is not finite`);
  }
  return Object.freeze({ kind: 'duration', ms });
}

/** True iff `c.value` is the formula variant. */
export function isFormulaCell(c: Cell): boolean {
  return typeof c.value === 'object' && c.value !== null && (c.value as { kind?: string }).kind === 'formula';
}

/** True iff `c.value` is the rich-text variant. */
export function isRichTextCell(c: Cell): boolean {
  return typeof c.value === 'object' && c.value !== null && (c.value as { kind?: string }).kind === 'rich-text';
}

/** Returns true iff the cell has no content. */
export function isEmptyCell(c: Cell): boolean {
  return c.value === null;
}
