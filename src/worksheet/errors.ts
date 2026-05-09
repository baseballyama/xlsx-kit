// Cell-level error / watch metadata. (cellWatches / ignoredErrors).
//
// `ignoredErrors` lets you tell Excel "don't flag these cells with the little
// green triangle for this class of validation". `cellWatches` records which
// cells the user has pinned in the Watch Window.
//
// Both round-trip via the worksheet `bodyExtras` passthrough already, but
// promoting them to first-class arrays gives the editor a clean API and the
// writer a stable position in the worksheet element order (between rowBreaks /
// colBreaks and the drawing block per ECMA-376 §18.3.1.94 cellWatches /
// §18.3.1.51 ignoredErrors).

import type { MultiCellRange } from './cell-range';

/**
 * One Cell-Watch entry. The Watch Window in Excel (Formulas → Watch Window)
 * shows live values for the cells listed here.
 */
export interface CellWatch {
  /** Single-cell reference like "Sheet1!$A$1" — kept verbatim. */
  ref: string;
}

/**
 * One ignored-error entry. Each Boolean flag corresponds to a class of Excel's
 * background validation; setting any of them to `true` suppresses the green
 * triangle for cells in `sqref`.
 */
export interface IgnoredError {
  /** Cells to which the suppressions apply. */
  sqref: MultiCellRange;
  /** "Formula evaluates to error" warning. */
  evalError?: boolean;
  /** "Date stored as 2-digit year" warning. */
  twoDigitTextYear?: boolean;
  /** "Number stored as text" warning. */
  numberStoredAsText?: boolean;
  /** "Inconsistent formula" warning. */
  formula?: boolean;
  /** "Formula omits cells" warning. */
  formulaRange?: boolean;
  /** "Unlocked cells containing formulas" warning. */
  unlockedFormula?: boolean;
  /** "Empty cells referenced" warning. */
  emptyCellReference?: boolean;
  /** "Data validation list error" warning. */
  listDataValidation?: boolean;
  /** "Inconsistent calculated column" warning (Excel Tables). */
  calculatedColumn?: boolean;
}

export const makeCellWatch = (ref: string): CellWatch => ({ ref });

export function makeIgnoredError(
  opts: Partial<IgnoredError> & { sqref: MultiCellRange },
): IgnoredError {
  return {
    sqref: opts.sqref,
    ...(opts.evalError !== undefined ? { evalError: opts.evalError } : {}),
    ...(opts.twoDigitTextYear !== undefined ? { twoDigitTextYear: opts.twoDigitTextYear } : {}),
    ...(opts.numberStoredAsText !== undefined ? { numberStoredAsText: opts.numberStoredAsText } : {}),
    ...(opts.formula !== undefined ? { formula: opts.formula } : {}),
    ...(opts.formulaRange !== undefined ? { formulaRange: opts.formulaRange } : {}),
    ...(opts.unlockedFormula !== undefined ? { unlockedFormula: opts.unlockedFormula } : {}),
    ...(opts.emptyCellReference !== undefined ? { emptyCellReference: opts.emptyCellReference } : {}),
    ...(opts.listDataValidation !== undefined ? { listDataValidation: opts.listDataValidation } : {}),
    ...(opts.calculatedColumn !== undefined ? { calculatedColumn: opts.calculatedColumn } : {}),
  };
}
