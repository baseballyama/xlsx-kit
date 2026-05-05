// Workbook-level <bookViews> typed model. Per ECMA-376 §18.2.30 and
// openpyxl/openpyxl/workbook/views.py.
//
// `<bookViews>` carries one or more `<workbookView>` entries. The
// first entry drives Excel's default tab strip (firstSheet / activeTab)
// and window position. Most workbooks have exactly one entry.

export type WorkbookViewVisibility = 'visible' | 'hidden' | 'veryHidden';

export interface WorkbookView {
  visibility?: WorkbookViewVisibility;
  minimized?: boolean;
  showHorizontalScroll?: boolean;
  showVerticalScroll?: boolean;
  showSheetTabs?: boolean;
  /** Window x position in screen pixels — Excel restores it when re-opening. */
  xWindow?: number;
  /** Window y position. */
  yWindow?: number;
  windowWidth?: number;
  windowHeight?: number;
  /** Width of the sheet tab strip relative to the horizontal scroll bar (0..1000, default 600). */
  tabRatio?: number;
  /** Index of the leftmost visible sheet tab (0-based). */
  firstSheet?: number;
  /** Index of the currently active sheet tab (0-based). */
  activeTab?: number;
  autoFilterDateGrouping?: boolean;
}

export const makeWorkbookView = (opts: WorkbookView = {}): WorkbookView => ({ ...opts });
