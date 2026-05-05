// AutoFilter. Per docs/plan/07-rich-features.md §4.
//
// **Stage 1**: ref + filterColumns where each entry is the
// `kind: 'filters'` variant — the value-list dropdown filter that
// covers >95% of real-world spreadsheets. customFilters / top10 /
// dynamicFilter / colorFilter / iconFilter / SortState are reserved
// for later iterations.

export type FilterColumn = {
  kind: 'filters';
  colId: number;
  /** Discrete values that pass the filter. Stored as strings to match the wire format. */
  values: string[];
  /** Whether blanks are visible. */
  blank?: boolean;
};

export interface AutoFilter {
  /** Excel range the filter covers (`"A1:E100"`). */
  ref: string;
  filterColumns: FilterColumn[];
}

export function makeAutoFilter(opts: { ref: string; filterColumns?: FilterColumn[] }): AutoFilter {
  return { ref: opts.ref, filterColumns: opts.filterColumns ?? [] };
}

export function makeFilterColumn(opts: {
  colId: number;
  values: ReadonlyArray<string>;
  blank?: boolean;
}): FilterColumn {
  return {
    kind: 'filters',
    colId: opts.colId,
    values: [...opts.values],
    ...(opts.blank !== undefined ? { blank: opts.blank } : {}),
  };
}

// ---- Worksheet ergonomic builders ---------------------------------------

import type { Worksheet } from './worksheet';

/** Add an AutoFilter dropdown header strip to the given range. */
export const addAutoFilter = (ws: Worksheet, ref: string): AutoFilter => {
  ws.autoFilter = makeAutoFilter({ ref });
  return ws.autoFilter;
};

/**
 * Add a value-list dropdown filter to a column inside the existing
 * AutoFilter range. `colId` is 0-based relative to the AutoFilter
 * left edge.
 */
export const addAutoFilterColumn = (
  ws: Worksheet,
  colId: number,
  values: ReadonlyArray<string>,
  opts: { blank?: boolean } = {},
): FilterColumn => {
  if (!ws.autoFilter) {
    throw new Error('addAutoFilterColumn: call addAutoFilter(ws, ref) first');
  }
  const fc = makeFilterColumn({ colId, values, ...(opts.blank !== undefined ? { blank: opts.blank } : {}) });
  ws.autoFilter.filterColumns.push(fc);
  return fc;
};

/** Drop the worksheet's AutoFilter entirely. */
export const removeAutoFilter = (ws: Worksheet): void => {
  delete (ws as { autoFilter?: AutoFilter }).autoFilter;
};
