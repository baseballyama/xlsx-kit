// Chartsheet model. Per docs/plan/08-charts-drawings.md §7.
//
// A chartsheet is a workbook child that holds a single chart instead
// of cells. It still gets a row in `<sheets>` (with the same r:id +
// sheetId machinery as a worksheet), but its part lives under
// `xl/chartsheets/sheetN.xml` and references a drawing via
// `<drawing r:id="..."/>` carrying an absoluteAnchor with the chart.

import type { Drawing } from '../drawing/drawing';
import type { HeaderFooter, PageMargins, PageSetup } from '../worksheet/page-setup';

/** `<sheetView>` inside `<sheetViews>`. */
export interface ChartsheetView {
  workbookViewId: number;
  tabSelected?: boolean;
  zoomScale?: number;
  zoomToFit?: boolean;
}

/** `<sheetPr>` chartsheet properties. */
export interface ChartsheetProperties {
  published?: boolean;
  codeName?: string;
  /** Tab color as RRGGBB (no alpha). */
  tabColorRgb?: string;
}

/** Subset of `<sheetProtection>` fields we round-trip on chartsheets. */
export interface ChartsheetProtection {
  /** Protect content (chart elements / text). */
  content?: boolean;
  /** Protect drawing objects (shapes / annotations). */
  objects?: boolean;
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
}

export interface Chartsheet {
  /** Display title shown in Excel's tab strip. Mirrors Worksheet.title for SheetRef compatibility. */
  title: string;
  views: ChartsheetView[];
  properties?: ChartsheetProperties;
  protection?: ChartsheetProtection;
  /**
   * Drawing payload — typically a single absoluteAnchor with a chart
   * graphicFrame. Mirrors Worksheet.drawing so the existing drawing-XML
   * helpers can be reused.
   */
  drawing?: Drawing;
  /** `<pageMargins>` — six required margins in inches. */
  pageMargins?: PageMargins;
  /** `<pageSetup>` — paper size / orientation / scale / fitToPage. */
  pageSetup?: PageSetup;
  /** `<headerFooter>` — odd/even/first header + footer mini-format strings. */
  headerFooter?: HeaderFooter;
  /** `<legacyDrawing r:id="…"/>` — VML drawing for comments / form-control overlay. */
  legacyDrawingRId?: string;
  /** `<legacyDrawingHF r:id="…"/>` — VML drawing for header/footer print background. */
  legacyDrawingHFRId?: string;
  /**
   * `<drawingHF r:id="…" lho="N" cho="N" lhe="N" che="N" lhf="N" chf="N"
   * rho="N" cho2="N" rhe="N" che2="N" rhf="N" chf2="N" lfo="N" cfo="N"
   * lfe="N" cfe="N" lff="N" cff="N" rfo="N" cfo2="N" rfe="N" cfe2="N"
   * rff="N" cff2="N"/>` — drawing slot + per-section image indices for
   * header/footer print backgrounds (DrawingML rather than VML).
   */
  drawingHF?: ChartsheetDrawingHF;
  /** `<picture r:id="…"/>` — chartsheet background image. */
  backgroundPictureRId?: string;
}

/**
 * `<drawingHF>` — per-section image-index map for the header/footer
 * drawing reference. Each `*o` / `*e` / `*f` attr is a 1-based image
 * number into the referenced drawing part. All optional.
 */
export interface ChartsheetDrawingHF {
  /** Required rels link to xl/drawings/drawingN.xml carrying the actual image refs. */
  rId: string;
  /** Left-header image index for odd pages. */
  lho?: number;
  cho?: number;
  rho?: number;
  /** Left-header image index for even pages. */
  lhe?: number;
  che?: number;
  rhe?: number;
  /** Left-header image index for the first page (when differentFirst). */
  lhf?: number;
  chf?: number;
  rhf?: number;
  /** Left-footer image index for odd pages. */
  lfo?: number;
  cfo?: number;
  rfo?: number;
  /** Left-footer image index for even pages. */
  lfe?: number;
  cfe?: number;
  rfe?: number;
  /** Left-footer image index for first page. */
  lff?: number;
  cff?: number;
  rff?: number;
}

export const makeChartsheet = (title: string): Chartsheet => ({
  title,
  views: [{ workbookViewId: 0, zoomToFit: true }],
});
