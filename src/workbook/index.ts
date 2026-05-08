// Public surface for the Workbook root model + sub-features (defined names,
// shared strings, protection, views, calc properties, workbook properties,
// file metadata, smart tags, function groups).

export type {
  CellSummary,
  SheetRef,
  SheetState,
  Workbook,
  WorkbookOverview,
  WorkbookSheetOverview,
  WorkbookStats,
} from './workbook';
export {
  addChartsheet,
  addWorksheet,
  countSheets,
  createWorkbook,
  createWorkbookFromCsv,
  createWorkbookFromCsvBundle,
  createWorkbookFromObjects,
  describeWorkbook,
  duplicateSheet,
  findCellInWorkbook,
  findCellsInWorkbook,
  findTable,
  getActiveSheet,
  getActiveSheetTitle,
  getAllCharts,
  getAllComments,
  getAllConditionalFormatting,
  getAllDataValidations,
  getAllHyperlinks,
  getAllImages,
  getAllMergedRanges,
  getAllTables,
  getCellAtAddress,
  getCellSummary,
  getChartsheet,
  getRangeValuesAtAddress,
  getSheet,
  getSheetByIndex,
  getSheetIndex,
  getSheetState,
  getSheetTitles,
  getValueAtAddress,
  getWorkbookAsCsvBundle,
  getWorkbookAsCsvRecord,
  getWorkbookAsHtmlRecord,
  getWorkbookAsMarkdownRecord,
  getWorkbookAsTextTableRecord,
  getWorkbookCellsByKind,
  getWorkbookStats,
  hasChartsheet,
  hasSheet,
  hasWorksheet,
  hideSheet,
  isActiveSheet,
  isValidSheetTitle,
  isWorkbookEmpty,
  iterAllCells,
  iterChartsheets,
  iterVisibleWorksheets,
  iterWorksheets,
  iterWorksheetsByState,
  jsonReplacer,
  jsonReviver,
  listChartsheets,
  listCustomXmlParts,
  listWorksheets,
  moveSheet,
  pickUniqueSheetTitle,
  removeSheet,
  renameSheet,
  replaceCellValuesInWorkbook,
  setActiveSheet,
  setCellAtAddress,
  setRangeValuesAtAddress,
  setSheetState,
  setSheetStates,
  sheetNames,
  showAllSheets,
  showSheet,
  swapSheets,
  validateSheetTitle,
  veryHideSheet,
} from './workbook';
export type { DefinedName, DefinedNameTarget } from './defined-names';
export {
  addDefinedName,
  addDefinedNameForRange,
  getDefinedName,
  getDefinedNameTarget,
  listDefinedNames,
  listPrintAreas,
  listPrintTitles,
  makeDefinedName,
  removeDefinedName,
  removeDefinedNames,
  renameDefinedName,
  setPrintArea,
  setPrintTitles,
} from './defined-names';
export type { WorkbookProtection } from './protection';
export {
  isWorkbookProtected,
  makeWorkbookProtection,
  protectWorkbook,
  unprotectWorkbook,
} from './protection';
export type { SharedStringEntry, SharedStringsTable } from './shared-strings';
export {
  addSharedString,
  getSharedStringAt,
  getSharedStringIndex,
  makeSharedStrings,
  sharedStringCount,
} from './shared-strings';
export type {
  CustomViewShowComments,
  CustomViewShowObjects,
  CustomWorkbookView,
  WorkbookView,
  WorkbookViewVisibility,
} from './views';
export {
  getActiveTab,
  getFirstSheet,
  makeCustomWorkbookView,
  makeWorkbookView,
  setActiveTab,
  setFirstSheet,
  setShowHorizontalScroll,
  setShowSheetTabs,
  setShowVerticalScroll,
  setTabRatio,
  setWorkbookMinimized,
  setWorkbookVisibility,
  setWorkbookWindow,
} from './views';
export type { CalcMode, CalcProperties, RefMode } from './calc-properties';
export {
  makeCalcProperties,
  setCalcMode,
  setCalcOnSave,
  setFullCalcOnLoad,
  setFullPrecision,
  setIterativeCalc,
} from './calc-properties';
export type {
  ShowObjectsMode,
  UpdateLinksMode,
  WorkbookProperties,
} from './workbook-properties';
export {
  makeWorkbookProperties,
  setDate1904,
  setFilterPrivacy,
  setUpdateLinksMode,
  setWorkbookCodeName,
} from './workbook-properties';
export type { FileVersion } from './file-version';
export { makeFileVersion } from './file-version';
export type { FileSharing } from './file-sharing';
export { makeFileSharing } from './file-sharing';
export type { FileRecoveryProperties } from './file-recovery';
export { makeFileRecoveryProperties } from './file-recovery';
export type { SmartTagProperties, SmartTagShowMode, SmartTagType } from './smart-tags';
export { makeSmartTagProperties, makeSmartTagType } from './smart-tags';
export type { FunctionGroup, FunctionGroups } from './function-groups';
export { makeFunctionGroup, makeFunctionGroups } from './function-groups';
