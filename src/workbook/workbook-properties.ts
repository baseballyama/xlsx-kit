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
