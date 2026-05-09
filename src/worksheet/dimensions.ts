// Column / row dimension metadata. (`columnDimensions` / `rowDimensions` Maps +
// `defaultColumnWidth` / `defaultRowHeight`) and the openpyxl reference at
// `worksheet/dimensions.py`.
//
// **Stage 1**: width / height / hidden / customWidth / customHeight / bestFit /
// outlineLevel + style fall-back. Multi-`<col>` runs collapse per-column
// entries when adjacent equal entries are persisted (writer side) — read keeps
// each column its own entry so user mutations don't surprise neighbours.

/**
 * Single-column-or-range dimension entry. Mirrors the OOXML `<col>` element.
 * The `min`/`max` pair always covers a contiguous run; the worksheet's
 * `columnDimensions` Map keys by `min` so per-column edits stay O(log n).
 */
export interface ColumnDimension {
  /** 1-based first column the entry covers (inclusive). */
  min: number;
  /** 1-based last column the entry covers (inclusive). */
  max: number;
  /** Width in Excel character units. */
  width?: number;
  /** Excel records `customWidth=1` whenever the user touched the slider. */
  customWidth?: boolean;
  /** Hidden columns get `hidden="1"`. */
  hidden?: boolean;
  /** "Best fit" auto-shrink semantics — Excel sets it for narrow columns. */
  bestFit?: boolean;
  /** Outline-level for grouping (0 = ungrouped). */
  outlineLevel?: number;
  /** Default cell xfId applied to empty cells in the column. */
  style?: number;
  /** Whether the column's outline is collapsed. */
  collapsed?: boolean;
}

/** Per-row metadata. Mirrors the OOXML `<row>` element's attributes. */
export interface RowDimension {
  /** Row height in points. */
  height?: number;
  /** Mirrors Excel's `customHeight="1"` flag. */
  customHeight?: boolean;
  hidden?: boolean;
  outlineLevel?: number;
  collapsed?: boolean;
  /** Default cell xfId applied to empty cells in the row. */
  style?: number;
}

/** Build a single-column ColumnDimension entry covering `col`. */
export function makeColumnDimension(
  col: number,
  opts: Partial<Omit<ColumnDimension, 'min' | 'max'>> = {},
): ColumnDimension {
  return {
    min: col,
    max: col,
    ...(opts.width !== undefined ? { width: opts.width } : {}),
    ...(opts.customWidth !== undefined ? { customWidth: opts.customWidth } : {}),
    ...(opts.hidden !== undefined ? { hidden: opts.hidden } : {}),
    ...(opts.bestFit !== undefined ? { bestFit: opts.bestFit } : {}),
    ...(opts.outlineLevel !== undefined ? { outlineLevel: opts.outlineLevel } : {}),
    ...(opts.style !== undefined ? { style: opts.style } : {}),
    ...(opts.collapsed !== undefined ? { collapsed: opts.collapsed } : {}),
  };
}

export function makeRowDimension(opts: Partial<RowDimension> = {}): RowDimension {
  return {
    ...(opts.height !== undefined ? { height: opts.height } : {}),
    ...(opts.customHeight !== undefined ? { customHeight: opts.customHeight } : {}),
    ...(opts.hidden !== undefined ? { hidden: opts.hidden } : {}),
    ...(opts.outlineLevel !== undefined ? { outlineLevel: opts.outlineLevel } : {}),
    ...(opts.collapsed !== undefined ? { collapsed: opts.collapsed } : {}),
    ...(opts.style !== undefined ? { style: opts.style } : {}),
  };
}
