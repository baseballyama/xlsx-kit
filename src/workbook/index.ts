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
  createWorkbook,
  describeWorkbook,
  getActiveSheet,
  getCellAtAddress,
  getCellSummary,
  getChartsheet,
  getSheet,
  getSheetState,
  getWorkbookCellsByKind,
  getWorkbookStats,
  iterWorksheets,
  listCustomXmlParts,
  moveSheet,
  removeSheet,
  renameSheet,
  setActiveSheet,
  setCellAtAddress,
  setSheetState,
  sheetNames,
} from './workbook';
export type { DefinedName, DefinedNameTarget } from './defined-names';
export {
  addDefinedName,
  getDefinedName,
  getDefinedNameTarget,
  listDefinedNames,
  makeDefinedName,
  removeDefinedName,
} from './defined-names';
export type { WorkbookProtection } from './protection';
export { makeWorkbookProtection } from './protection';
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
export { makeCustomWorkbookView, makeWorkbookView } from './views';
export type { CalcMode, CalcProperties, RefMode } from './calc-properties';
export { makeCalcProperties } from './calc-properties';
export type {
  ShowObjectsMode,
  UpdateLinksMode,
  WorkbookProperties,
} from './workbook-properties';
export { makeWorkbookProperties } from './workbook-properties';
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
