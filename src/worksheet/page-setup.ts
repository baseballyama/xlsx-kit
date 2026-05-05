// Page-setup typed model. Per docs/plan/13-full-excel-coverage.md §B6.
//
// Promotes <printOptions> / <pageMargins> / <pageSetup> / <headerFooter>
// from the bodyExtras passthrough into typed Worksheet fields with
// round-trip readers / writers. Mirrors openpyxl/openpyxl/worksheet/
// page.py + header_footer.py.

export interface PrintOptions {
  /** Center the printed sheet horizontally on the page. */
  horizontalCentered?: boolean;
  /** Center the printed sheet vertically on the page. */
  verticalCentered?: boolean;
  /** Print row + column headings (the A B C / 1 2 3 strips). */
  headings?: boolean;
  /** Print sheet gridlines. */
  gridLines?: boolean;
  /** Mirrors a quirky Excel companion flag for `gridLines`. */
  gridLinesSet?: boolean;
}

/** Page margins in inches. ECMA-376 §18.3.1.62. All six fields are required when the element is present. */
export interface PageMargins {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

export type PageOrientation = 'default' | 'portrait' | 'landscape';
export type PageOrder = 'downThenOver' | 'overThenDown';
export type CellCommentMode = 'none' | 'asDisplayed' | 'atEnd';
export type PrintErrorMode = 'displayed' | 'blank' | 'dash' | 'NA';

export interface PageSetup {
  paperSize?: number;
  scale?: number;
  firstPageNumber?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  pageOrder?: PageOrder;
  orientation?: PageOrientation;
  usePrinterDefaults?: boolean;
  blackAndWhite?: boolean;
  draft?: boolean;
  cellComments?: CellCommentMode;
  useFirstPageNumber?: boolean;
  errors?: PrintErrorMode;
  horizontalDpi?: number;
  verticalDpi?: number;
  copies?: number;
  /** Optional `r:id` referencing an external printerSettings part — round-tripped verbatim. */
  rId?: string;
  /** Paper height (UniversalMeasure, e.g. "297mm"). */
  paperHeight?: string;
  /** Paper width (UniversalMeasure). */
  paperWidth?: string;
}

export interface HeaderFooter {
  differentFirst?: boolean;
  differentOddEven?: boolean;
  /** Mirror Excel's "scale header/footer with document" toggle. Default true. */
  scaleWithDoc?: boolean;
  /** Mirror Excel's "align header/footer with margins" toggle. Default true. */
  alignWithMargins?: boolean;
  /**
   * Mini-format string. Excel uses `&L` / `&C` / `&R` to split the
   * three sections, plus codes like `&P` (page number), `&N` (page
   * count), `&F` (file name), `&A` (sheet name), `&D` / `&T` (date /
   * time). We round-trip the literal text — no parsing into sections.
   */
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
}

export const makePageMargins = (opts: Partial<PageMargins> = {}): PageMargins => ({
  left: opts.left ?? 0.75,
  right: opts.right ?? 0.75,
  top: opts.top ?? 1,
  bottom: opts.bottom ?? 1,
  header: opts.header ?? 0.5,
  footer: opts.footer ?? 0.5,
});

export const makePrintOptions = (opts: PrintOptions = {}): PrintOptions => ({ ...opts });

export const makePageSetup = (opts: PageSetup = {}): PageSetup => ({ ...opts });

export const makeHeaderFooter = (opts: HeaderFooter = {}): HeaderFooter => ({ ...opts });

/**
 * One manual page break. `id` is the row (for rowBreaks) or column
 * (for colBreaks) index where the break sits; `min`/`max` constrain the
 * orthogonal range Excel honours; `man=true` means a user-placed break
 * (default true). `pt` indicates a "pivot table" break — rare.
 */
export interface PageBreak {
  id?: number;
  min?: number;
  max?: number;
  man?: boolean;
  pt?: boolean;
}

export const makePageBreak = (opts: PageBreak = {}): PageBreak => ({ ...opts });
