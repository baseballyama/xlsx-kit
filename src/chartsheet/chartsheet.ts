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
}

export const makeChartsheet = (title: string): Chartsheet => ({
  title,
  views: [{ workbookViewId: 0, zoomToFit: true }],
});
