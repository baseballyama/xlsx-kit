// Worksheet-level <sortState>. Per ECMA-376 §18.3.1.92.
//
// Excel persists the last sort the user applied so re-opening the file
// shows the rows in the same order. The element carries one or more
// <sortCondition> entries describing the sort key columns (or rows).

export type SortBy = 'value' | 'cellColor' | 'fontColor' | 'icon';
export type SortMethod = 'stroke' | 'pinYin';
export type SortIconSet =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '4Arrows'
  | '4ArrowsGray'
  | '4RedToBlack'
  | '4Rating'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Rating'
  | '5Quarters';

export interface SortCondition {
  /** Column or row that drives this sort key. */
  ref: string;
  descending?: boolean;
  sortBy?: SortBy;
  /** Reference to a custom-list defined name. */
  customList?: string;
  /** Differential-style index (`<dxf>` slot in the stylesheet). */
  dxfId?: number;
  iconSet?: SortIconSet;
  iconId?: number;
}

export interface SortState {
  /** Range the sort applies to (`A1:D20`). */
  ref: string;
  conditions: SortCondition[];
  /** Sort columns instead of rows (rare). */
  columnSort?: boolean;
  caseSensitive?: boolean;
  sortMethod?: SortMethod;
}

export const makeSortCondition = (opts: SortCondition): SortCondition => ({ ...opts });

export const makeSortState = (opts: Partial<SortState> & { ref: string }): SortState => ({
  ref: opts.ref,
  conditions: opts.conditions?.slice() ?? [],
  ...(opts.columnSort !== undefined ? { columnSort: opts.columnSort } : {}),
  ...(opts.caseSensitive !== undefined ? { caseSensitive: opts.caseSensitive } : {}),
  ...(opts.sortMethod !== undefined ? { sortMethod: opts.sortMethod } : {}),
});
