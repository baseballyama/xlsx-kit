// Data validations. Per docs/plan/07-rich-features.md §5.
//
// A DataValidation entry attaches a constraint (one of seven type
// kinds, optional operator, two formula slots) to a sqref-style
// MultiCellRange. Stage-1 maps every OOXML attribute we have a use for;
// imeMode + numeric/value clamps land later when phase 7's Asian-locale
// support catches up.

import type { MultiCellRange } from './cell-range';

export type DataValidationType = 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';
export type DataValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'greaterThanOrEqual'
  | 'lessThan'
  | 'lessThanOrEqual';
export type DataValidationErrorStyle = 'stop' | 'warning' | 'information';

export interface DataValidation {
  /** Constraint kind. `'list'` is the dropdown form (formula1 = comma list or range). */
  type: DataValidationType;
  /** Comparison operator — only meaningful for whole/decimal/date/time/textLength. */
  operator?: DataValidationOperator;
  /** Lower-bound formula or list source. Always required for list/whole/decimal/date/time. */
  formula1?: string;
  /** Upper-bound formula. Used by between / notBetween only. */
  formula2?: string;
  /** Allow empty cells. */
  allowBlank?: boolean;
  /** Show the input message popover when the cell is selected. */
  showInputMessage?: boolean;
  /** Show the error message popover on invalid entry. */
  showErrorMessage?: boolean;
  errorTitle?: string;
  error?: string;
  errorStyle?: DataValidationErrorStyle;
  promptTitle?: string;
  prompt?: string;
  /** Excel inverts this attribute: `showDropDown="1"` actually *hides* the dropdown arrow on list-type validators. We mirror the wire form. */
  showDropDown?: boolean;
  /** Apply-to range (sqref). */
  sqref: MultiCellRange;
}

export function makeDataValidation(
  opts: Partial<DataValidation> & { type: DataValidationType; sqref: MultiCellRange },
): DataValidation {
  return {
    type: opts.type,
    sqref: opts.sqref,
    ...(opts.operator !== undefined ? { operator: opts.operator } : {}),
    ...(opts.formula1 !== undefined ? { formula1: opts.formula1 } : {}),
    ...(opts.formula2 !== undefined ? { formula2: opts.formula2 } : {}),
    ...(opts.allowBlank !== undefined ? { allowBlank: opts.allowBlank } : {}),
    ...(opts.showInputMessage !== undefined ? { showInputMessage: opts.showInputMessage } : {}),
    ...(opts.showErrorMessage !== undefined ? { showErrorMessage: opts.showErrorMessage } : {}),
    ...(opts.errorTitle !== undefined ? { errorTitle: opts.errorTitle } : {}),
    ...(opts.error !== undefined ? { error: opts.error } : {}),
    ...(opts.errorStyle !== undefined ? { errorStyle: opts.errorStyle } : {}),
    ...(opts.promptTitle !== undefined ? { promptTitle: opts.promptTitle } : {}),
    ...(opts.prompt !== undefined ? { prompt: opts.prompt } : {}),
    ...(opts.showDropDown !== undefined ? { showDropDown: opts.showDropDown } : {}),
  };
}
