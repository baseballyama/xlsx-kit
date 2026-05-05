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

export type CustomViewShowComments = 'commNone' | 'commIndicator' | 'commIndAndComment';
export type CustomViewShowObjects = 'all' | 'placeholders' | 'none';

export interface CustomWorkbookView {
  name: string;
  guid: string;
  /** Window width in screen pixels — required when the element is present. */
  windowWidth: number;
  windowHeight: number;
  /** Index (0-based) of the sheet active in this saved view. */
  activeSheetId: number;
  autoUpdate?: boolean;
  /** Auto-merge interval in minutes (for shared workbooks). */
  mergeInterval?: number;
  changesSavedWin?: boolean;
  onlySync?: boolean;
  personalView?: boolean;
  includePrintSettings?: boolean;
  includeHiddenRowCol?: boolean;
  maximized?: boolean;
  minimized?: boolean;
  showHorizontalScroll?: boolean;
  showVerticalScroll?: boolean;
  showSheetTabs?: boolean;
  xWindow?: number;
  yWindow?: number;
  tabRatio?: number;
  showFormulaBar?: boolean;
  showStatusbar?: boolean;
  showComments?: CustomViewShowComments;
  showObjects?: CustomViewShowObjects;
}

export const makeCustomWorkbookView = (
  opts: Pick<CustomWorkbookView, 'name' | 'guid' | 'windowWidth' | 'windowHeight' | 'activeSheetId'> &
    Partial<CustomWorkbookView>,
): CustomWorkbookView => ({ ...opts });
