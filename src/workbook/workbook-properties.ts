// Workbook-level <workbookPr>. Per ECMA-376 §18.2.28.
//
// Mirrors openpyxl/openpyxl/workbook/properties.py WorkbookProperties.
// `date1904` is already lifted via wb.date1904; everything else lived
// in workbookXmlExtras passthrough until now. This typed shell promotes
// all 19 attrs into the modeled workbook so consumers can edit them
// without rebuilding the XmlNode.

export type ShowObjectsMode = 'all' | 'placeholders' | 'none';
export type UpdateLinksMode = 'userSet' | 'never' | 'always';

export interface WorkbookProperties {
  /**
   * Mac 1904 epoch flag. Mirrored from `Workbook.date1904` — the
   * canonical source. Setting it here has no effect on the cell-serial
   * conversion path, which keys off the top-level field.
   */
  date1904?: boolean;
  /** Excel 5/95 ↔ 2007 compatibility hint. */
  dateCompatibility?: boolean;
  showObjects?: ShowObjectsMode;
  showBorderUnselectedTables?: boolean;
  filterPrivacy?: boolean;
  promptedSolutions?: boolean;
  showInkAnnotation?: boolean;
  backupFile?: boolean;
  /** Cache external-link values when saving. */
  saveExternalLinkValues?: boolean;
  /** "Update remote links" prompt mode. */
  updateLinks?: UpdateLinksMode;
  /** VBA codeName for the workbook (e.g. "ThisWorkbook" / "ЭтаКнига"). */
  codeName?: string;
  hidePivotFieldList?: boolean;
  showPivotChartFilter?: boolean;
  allowRefreshQuery?: boolean;
  publishItems?: boolean;
  checkCompatibility?: boolean;
  autoCompressPictures?: boolean;
  refreshAllConnections?: boolean;
  /** Theme schema version (Excel 2007 = 124226, 2013+ = 153222). */
  defaultThemeVersion?: number;
}

export const makeWorkbookProperties = (opts: WorkbookProperties = {}): WorkbookProperties => ({ ...opts });

// ---- Workbook ergonomic helpers ----------------------------------------

import type { Workbook } from './workbook';

const ensureWorkbookProperties = (wb: Workbook): WorkbookProperties => {
  if (!wb.workbookProperties) wb.workbookProperties = {};
  return wb.workbookProperties;
};

/**
 * Set the workbook-level VBA codeName ("ThisWorkbook" by default in
 * Excel; localised forms like "ЭтаКнига" round-trip too). Empty string
 * is allowed — Excel writes it that way for codename-stripped files.
 */
export const setWorkbookCodeName = (wb: Workbook, codeName: string): void => {
  ensureWorkbookProperties(wb).codeName = codeName;
};

/**
 * Toggle the Mac 1904 epoch. The canonical flag is `wb.date1904`
 * (drives cell-serial conversion); this helper writes both the
 * canonical field and the mirror on `workbookProperties` so a save
 * emits a consistent `<workbookPr date1904="…">` attribute.
 */
export const setDate1904 = (wb: Workbook, on: boolean): void => {
  wb.date1904 = on;
  ensureWorkbookProperties(wb).date1904 = on;
};

/**
 * Set the "Update remote links" prompt mode. `'userSet'` keeps
 * Excel's per-user preference; `'never'` disables the prompt;
 * `'always'` forces it. Mirrors the Trust Center "External Content"
 * dropdown.
 */
export const setUpdateLinksMode = (wb: Workbook, mode: UpdateLinksMode): void => {
  ensureWorkbookProperties(wb).updateLinks = mode;
};

/** Toggle the "filterPrivacy" hint Excel writes to indicate filter contents may be sensitive. */
export const setFilterPrivacy = (wb: Workbook, on: boolean): void => {
  ensureWorkbookProperties(wb).filterPrivacy = on;
};
