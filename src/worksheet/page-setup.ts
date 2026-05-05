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

// ---- Worksheet ergonomic helpers ----------------------------------------
// Operate on a Worksheet directly so callers don't have to allocate the
// individual typed records up front.

import type { Worksheet } from './worksheet';

const ensurePageSetup = (ws: Worksheet): PageSetup => {
  if (!ws.pageSetup) ws.pageSetup = {};
  return ws.pageSetup;
};

const ensureHeaderFooter = (ws: Worksheet): HeaderFooter => {
  if (!ws.headerFooter) ws.headerFooter = {};
  return ws.headerFooter;
};

/** Set page orientation on `ws.pageSetup` (allocates if missing). */
export const setPageOrientation = (ws: Worksheet, orientation: PageOrientation): void => {
  ensurePageSetup(ws).orientation = orientation;
};

/** Set paper size code (Excel uses ECMA-376 §3.3 paper-size enums; 1=Letter, 9=A4 etc.). */
export const setPaperSize = (ws: Worksheet, paperSize: number): void => {
  ensurePageSetup(ws).paperSize = paperSize;
};

/** Set the print scale percentage (10..400). */
export const setPrintScale = (ws: Worksheet, scale: number): void => {
  ensurePageSetup(ws).scale = scale;
};

/** Set fitToWidth + fitToHeight (Excel "Fit to N pages wide × M tall" UI). */
export const setFitToPage = (ws: Worksheet, opts: { width?: number; height?: number }): void => {
  const ps = ensurePageSetup(ws);
  if (opts.width !== undefined) ps.fitToWidth = opts.width;
  if (opts.height !== undefined) ps.fitToHeight = opts.height;
};

/** Replace ws.pageMargins with the provided values (uses Excel defaults for missing axes). */
export const setPageMargins = (ws: Worksheet, opts: Partial<PageMargins> = {}): void => {
  ws.pageMargins = {
    left: opts.left ?? 0.75,
    right: opts.right ?? 0.75,
    top: opts.top ?? 1,
    bottom: opts.bottom ?? 1,
    header: opts.header ?? 0.5,
    footer: opts.footer ?? 0.5,
  };
};

export type HeaderFooterSection = 'odd' | 'even' | 'first';

/** Set the header text for a given section. Excel uses `&L` / `&C` / `&R` codes inside the string. */
export const setHeader = (ws: Worksheet, section: HeaderFooterSection, text: string): void => {
  const hf = ensureHeaderFooter(ws);
  if (section === 'odd') hf.oddHeader = text;
  else if (section === 'even') {
    hf.evenHeader = text;
    hf.differentOddEven = true;
  } else {
    hf.firstHeader = text;
    hf.differentFirst = true;
  }
};

/** Set the footer text for a given section. */
export const setFooter = (ws: Worksheet, section: HeaderFooterSection, text: string): void => {
  const hf = ensureHeaderFooter(ws);
  if (section === 'odd') hf.oddFooter = text;
  else if (section === 'even') {
    hf.evenFooter = text;
    hf.differentOddEven = true;
  } else {
    hf.firstFooter = text;
    hf.differentFirst = true;
  }
};

/**
 * Excel's reserved header / footer code tokens. Drop these into the
 * left / center / right text inputs of {@link buildHeaderFooterText}
 * (or directly into a setHeader / setFooter string) to render dynamic
 * values at print time.
 */
export const HEADER_FOOTER_CODES = Object.freeze({
  /** Current page number. */
  pageNumber: '&P',
  /** Total number of pages. */
  pageCount: '&N',
  /** Print date. */
  date: '&D',
  /** Print time. */
  time: '&T',
  /** File path + name. */
  filePath: '&Z&F',
  /** File name only. */
  fileName: '&F',
  /** Sheet name. */
  sheetName: '&A',
  /** Embedded image (Excel inserts via "Insert Picture" — `&G` is the placeholder). */
  picture: '&G',
});

/**
 * Build a header / footer string from optional left / center / right
 * fragments using Excel's `&L` / `&C` / `&R` markers. An empty fragment
 * is omitted (no marker emitted) so a center-only header doesn't leave
 * a stray `&L` prefix. Returns `''` when all three fragments are
 * undefined.
 */
export const buildHeaderFooterText = (
  parts: { left?: string; center?: string; right?: string },
): string => {
  let out = '';
  if (parts.left !== undefined) out += `&L${parts.left}`;
  if (parts.center !== undefined) out += `&C${parts.center}`;
  if (parts.right !== undefined) out += `&R${parts.right}`;
  return out;
};

/**
 * Set a header by left / center / right parts. `section` defaults
 * to `'odd'` (the standard pages); pass `'first'` or `'even'` to
 * target the alternate sections (Excel auto-flips the corresponding
 * differentOddEven / differentFirst flag).
 */
export const setHeaderText = (
  ws: Worksheet,
  parts: { left?: string; center?: string; right?: string },
  section: HeaderFooterSection = 'odd',
): void => {
  setHeader(ws, section, buildHeaderFooterText(parts));
};

/** Same shape as {@link setHeaderText} but writes the corresponding footer slot. */
export const setFooterText = (
  ws: Worksheet,
  parts: { left?: string; center?: string; right?: string },
  section: HeaderFooterSection = 'odd',
): void => {
  setFooter(ws, section, buildHeaderFooterText(parts));
};

/** Push a manual horizontal page break above the given row (1-based). Defaults to `man=true`. */
export const addRowBreak = (ws: Worksheet, row: number): PageBreak => {
  const brk: PageBreak = { id: row, man: true, max: 16383 };
  ws.rowBreaks.push(brk);
  return brk;
};

/** Push a manual vertical page break to the left of the given column (1-based). Defaults to `man=true`. */
export const addColBreak = (ws: Worksheet, col: number): PageBreak => {
  const brk: PageBreak = { id: col, man: true, max: 1048575 };
  ws.colBreaks.push(brk);
  return brk;
};

const ensurePrintOptions = (ws: Worksheet): PrintOptions => {
  if (!ws.printOptions) ws.printOptions = {};
  return ws.printOptions;
};

/** Toggle "Print gridlines". Mirrors Excel's "Page Layout → Sheet Options → Gridlines: Print". */
export const setPrintGridLines = (ws: Worksheet, on: boolean): void => {
  const po = ensurePrintOptions(ws);
  po.gridLines = on;
  // Excel pairs gridLines with the gridLinesSet companion flag.
  po.gridLinesSet = on;
};

/** Toggle "Print row and column headings" (the A B C / 1 2 3 strips on the printed page). */
export const setPrintHeadings = (ws: Worksheet, on: boolean): void => {
  ensurePrintOptions(ws).headings = on;
};

/**
 * Toggle horizontal / vertical centering on the printed page. Pass
 * either field to leave the other untouched.
 */
export const setPrintCentered = (
  ws: Worksheet,
  opts: { horizontal?: boolean; vertical?: boolean },
): void => {
  const po = ensurePrintOptions(ws);
  if (opts.horizontal !== undefined) po.horizontalCentered = opts.horizontal;
  if (opts.vertical !== undefined) po.verticalCentered = opts.vertical;
};

