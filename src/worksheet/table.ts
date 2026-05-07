// Excel Table object (xl/tables/tableN.xml). Per
// docs/plan/07-rich-features.md §3.
//
// Tables ride on top of a worksheet range, give it a name + structured
// column references, and own their own AutoFilter. Each table sits in
// a separate part — the worksheet only carries a `<tableParts>` block
// pointing at the workbook-rels rId. Stage-1 covers the table shell +
// columns + styleInfo + autoFilter; sortState / totals row formulas /
// calculated column formulas / xml extlst are reserved for later.

import type { AutoFilter } from './auto-filter';

export interface TableColumn {
  /** 1-based column id (per-table). */
  id: number;
  /** Header name. */
  name: string;
  /** Totals-row aggregation function. */
  totalsRowFunction?: 'sum' | 'min' | 'max' | 'count' | 'countNums' | 'average' | 'stdDev' | 'var' | 'custom';
  /** Override label for the totals row. */
  totalsRowLabel?: string;
  /** Custom totals-row formula text. */
  totalsRowFormula?: string;
  /** Calculated-column formula. */
  calculatedColumnFormula?: string;
}

export interface TableStyleInfo {
  /** Built-in style name (TableStyleMedium2, etc) or custom. */
  name?: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
}

export interface TableDefinition {
  /** Workbook-unique id (`<table id="N">`). */
  id: number;
  /** Workbook-unique displayName — Excel surfaces this in formulas. */
  displayName: string;
  /** Optional friendly name; usually matches `displayName`. */
  name?: string;
  /** Range covered by the table, e.g. "A1:E10". */
  ref: string;
  /** Number of header rows. Defaults to 1; 0 means a header-less table. */
  headerRowCount?: number;
  /** Number of totals rows. */
  totalsRowCount?: number;
  /** Whether the totals row is currently visible. */
  totalsRowShown?: boolean;
  styleInfo?: TableStyleInfo;
  columns: TableColumn[];
  autoFilter?: AutoFilter;
  /** Worksheet-rels rId — populated on read; the writer assigns its own. */
  rId?: string;
}

export function makeTableColumn(opts: { id: number; name: string }): TableColumn {
  return { id: opts.id, name: opts.name };
}

export function makeTableDefinition(opts: {
  id: number;
  displayName: string;
  ref: string;
  name?: string;
  columns?: TableColumn[];
  headerRowCount?: number;
  totalsRowCount?: number;
  totalsRowShown?: boolean;
  styleInfo?: TableStyleInfo;
  autoFilter?: AutoFilter;
}): TableDefinition {
  return {
    id: opts.id,
    displayName: opts.displayName,
    ref: opts.ref,
    columns: opts.columns ?? [],
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    ...(opts.headerRowCount !== undefined ? { headerRowCount: opts.headerRowCount } : {}),
    ...(opts.totalsRowCount !== undefined ? { totalsRowCount: opts.totalsRowCount } : {}),
    ...(opts.totalsRowShown !== undefined ? { totalsRowShown: opts.totalsRowShown } : {}),
    ...(opts.styleInfo ? { styleInfo: opts.styleInfo } : {}),
    ...(opts.autoFilter ? { autoFilter: opts.autoFilter } : {}),
  };
}

// ---- Worksheet ergonomic builders ---------------------------------------

import type { CellValue } from '../cell/cell';
import { boundariesToRangeString } from '../utils/coordinate';
import { OpenXmlSchemaError } from '../utils/exceptions';
import type { Workbook } from '../workbook/workbook';
import type { Worksheet } from './worksheet';
import { writeRangeFromObjects } from './worksheet';

const nextTableId = (wb: Workbook): number => {
  let max = 0;
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    for (const t of ref.sheet.tables) {
      if (t.id > max) max = t.id;
    }
  }
  return max + 1;
};

/**
 * High-level wrapper that builds a TableDefinition + pushes it onto
 * `ws.tables` in one call. Auto-assigns the workbook-unique `id`,
 * derives `displayName` from the supplied `name`, and constructs
 * `TableColumn` records (1-based ids) from a string-array shorthand.
 *
 * Pass `style` for one-arg style selection; for finer-grained control
 * (showFirstColumn / showLastColumn / row-stripes / column-stripes)
 * pass the full `styleInfo` instead.
 */
export const addExcelTable = (
  wb: Workbook,
  ws: Worksheet,
  opts: {
    name: string;
    ref: string;
    columns: ReadonlyArray<string | TableColumn>;
    /** Built-in style name shortcut (e.g. 'TableStyleMedium2'). Ignored if `styleInfo` is supplied. */
    style?: string;
    styleInfo?: TableStyleInfo;
    headerRowCount?: number;
    totalsRowCount?: number;
    totalsRowShown?: boolean;
    autoFilter?: AutoFilter;
    /** Override displayName (defaults to `name`). */
    displayName?: string;
  },
): TableDefinition => {
  const cols: TableColumn[] = opts.columns.map((c, i): TableColumn =>
    typeof c === 'string' ? { id: i + 1, name: c } : c,
  );
  const styleInfo: TableStyleInfo | undefined =
    opts.styleInfo ??
    (opts.style !== undefined ? { name: opts.style, showRowStripes: true, showColumnStripes: false } : undefined);
  const def: TableDefinition = makeTableDefinition({
    id: nextTableId(wb),
    displayName: opts.displayName ?? opts.name,
    name: opts.name,
    ref: opts.ref,
    columns: cols,
    ...(opts.headerRowCount !== undefined ? { headerRowCount: opts.headerRowCount } : {}),
    ...(opts.totalsRowCount !== undefined ? { totalsRowCount: opts.totalsRowCount } : {}),
    ...(opts.totalsRowShown !== undefined ? { totalsRowShown: opts.totalsRowShown } : {}),
    ...(styleInfo ? { styleInfo } : {}),
    ...(opts.autoFilter ? { autoFilter: opts.autoFilter } : {}),
  });
  ws.tables.push(def);
  return def;
};

/**
 * Convenience: write a `Record[]` block to the sheet via
 * {@link writeRangeFromObjects}, then register an Excel table over
 * the bounding-box. Combines the two-step "write the values, then
 * declare the table" pattern into a single call.
 *
 * Throws on empty input — there's no meaningful zero-row table.
 *
 * Returns the registered {@link TableDefinition} (the same object
 * pushed onto `ws.tables`). Use `opts.headers` to pin column order;
 * the columns array on the resulting table mirrors that order.
 */
export const addTableFromObjects = (
  wb: Workbook,
  ws: Worksheet,
  opts: {
    name: string;
    startRef: string;
    objects: ReadonlyArray<Record<string, CellValue | null | undefined>>;
    headers?: ReadonlyArray<string>;
    style?: string;
    styleInfo?: TableStyleInfo;
    displayName?: string;
  },
): TableDefinition => {
  if (opts.objects.length === 0) {
    throw new OpenXmlSchemaError('addTableFromObjects: objects array must be non-empty');
  }
  const writeOpts: { headers?: ReadonlyArray<string> } = opts.headers ? { headers: opts.headers } : {};
  const bounds = writeRangeFromObjects(ws, opts.startRef, opts.objects, writeOpts);
  if (!bounds) {
    throw new OpenXmlSchemaError('addTableFromObjects: writeRangeFromObjects returned no bounds');
  }
  // Recompute the canonical header order from the same logic
  // writeRangeFromObjects used (so the table's columns line up).
  let headers: string[];
  if (opts.headers) headers = [...opts.headers];
  else {
    const seen = new Set<string>();
    headers = [];
    for (const o of opts.objects) {
      for (const k of Object.keys(o)) {
        if (!seen.has(k)) {
          seen.add(k);
          headers.push(k);
        }
      }
    }
  }
  const ref = boundariesToRangeString(bounds);
  return addExcelTable(wb, ws, {
    name: opts.name,
    ref,
    columns: headers,
    ...(opts.style !== undefined ? { style: opts.style } : {}),
    ...(opts.styleInfo ? { styleInfo: opts.styleInfo } : {}),
    ...(opts.displayName !== undefined ? { displayName: opts.displayName } : {}),
  });
};
