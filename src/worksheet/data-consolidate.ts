// Worksheet-level <dataConsolidate>. Used by Excel's
// Data → Consolidate dialog to join multiple ranges into one summary
// table. Per ECMA-376 §18.3.1.20, §18.3.1.22 (dataRef).

export type DataConsolidateFunction =
  | 'average'
  | 'count'
  | 'countNums'
  | 'max'
  | 'min'
  | 'product'
  | 'stdDev'
  | 'stdDevp'
  | 'sum'
  | 'var'
  | 'varp';

export interface DataReference {
  /** Optional name of the source range (defined-name reference). */
  name?: string;
  /** External range ref like "Sheet1!$A$1:$B$10" — required when no `rId`. */
  ref?: string;
  /** Optional friendly sheet name for display. */
  sheet?: string;
  /** rels rId pointing at an external workbook part — round-tripped verbatim. */
  rId?: string;
}

export interface DataConsolidate {
  /** Aggregation function applied to overlapping cells. Default `sum`. */
  function?: DataConsolidateFunction;
  /** Use top-row labels as category keys. */
  topLabels?: boolean;
  /** Use left-column labels as category keys. */
  leftLabels?: boolean;
  /** "Create links to source data" (Excel's checkbox). */
  link?: boolean;
  /** Optional `<dataRefs>` list (one entry per source range). */
  dataRefs?: DataReference[];
  /**
   * `startLabels` was added in a later schema revision — round-tripped
   * verbatim when present.
   */
  startLabels?: string;
}

export const makeDataConsolidate = (opts: DataConsolidate = {}): DataConsolidate => {
  const out: DataConsolidate = {};
  if (opts.function !== undefined) out.function = opts.function;
  if (opts.topLabels !== undefined) out.topLabels = opts.topLabels;
  if (opts.leftLabels !== undefined) out.leftLabels = opts.leftLabels;
  if (opts.link !== undefined) out.link = opts.link;
  if (opts.dataRefs !== undefined) out.dataRefs = opts.dataRefs.map((r) => ({ ...r }));
  if (opts.startLabels !== undefined) out.startLabels = opts.startLabels;
  return out;
};
