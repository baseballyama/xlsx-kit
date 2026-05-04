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
