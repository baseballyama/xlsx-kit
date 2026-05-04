// Worksheet hyperlinks. Per docs/plan/07-rich-features.md §2.
//
// External URLs land in `xl/worksheets/_rels/sheetN.xml.rels` (one rel
// per Hyperlink entry, type "...relationships/hyperlink",
// TargetMode="External"). Internal jumps (`#'Sheet 2'!A1`) live entirely
// in the `<hyperlink location="..."/>` attribute and don't need a rel.

export interface Hyperlink {
  /** Cell or range the hyperlink covers — "A1" or "A1:B5". */
  ref: string;
  /** External URL or relative target path. Mutually exclusive with `location`-only links. */
  target?: string;
  /** Anchor inside the workbook (e.g. `'Sheet 2'!A1`). */
  location?: string;
  /** Tooltip shown on hover. */
  tooltip?: string;
  /** Visible link text — typically falls back to the referenced cell value. */
  display?: string;
  /** Worksheet-rels rId. Populated on read; assigned by the writer when missing. */
  rId?: string;
}

export function makeHyperlink(opts: Partial<Hyperlink> & { ref: string }): Hyperlink {
  if (opts.ref === undefined || opts.ref.length === 0) {
    throw new Error('Hyperlink: ref is required');
  }
  return {
    ref: opts.ref,
    ...(opts.target !== undefined ? { target: opts.target } : {}),
    ...(opts.location !== undefined ? { location: opts.location } : {}),
    ...(opts.tooltip !== undefined ? { tooltip: opts.tooltip } : {}),
    ...(opts.display !== undefined ? { display: opts.display } : {}),
    ...(opts.rId !== undefined ? { rId: opts.rId } : {}),
  };
}
