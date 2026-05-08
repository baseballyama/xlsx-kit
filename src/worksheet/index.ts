// Public surface for the Worksheet model + cell-range / dimensions / views /
// comments / hyperlinks / data-validations / conditional-formatting /
// auto-filter / tables / page-setup / protection / errors / smart-tags /
// ole-objects / sort-state / scenarios / data-consolidate / web-publish /
// phonetic / protected-ranges / properties / custom-sheet-views.

export type {
  CellsByKindCounts,
  IterRowsOptions,
  Worksheet,
} from './worksheet';
export {
  addCellWatch,
  addConditionalFormatting,
  addDataValidation,
  addIgnoredError,
  addTable,
  appendRow,
  appendRows,
  applyToRange,
  autofitColumns,
  clearAllCells,
  clearRange,
  collapseColumnGroup,
  collapseRowGroup,
  copyRange,
  countCells,
  countCellsByKind,
  deleteCell,
  expandColumnGroup,
  expandRowGroup,
  findCells,
  freezePanes,
  getAutoFilter,
  getCell,
  getCellByCoord,
  getCellsInColumn,
  getCellsInRange,
  getCellsInRow,
  getColumnDimension,
  getDataExtent,
  getFreezePanes,
  getMaxCol,
  getMaxRow,
  getMergedCells,
  getMergedRangeAt,
  getNonEmptyCellCount,
  getPopulatedColumnIndices,
  getPopulatedRowIndices,
  getRangeValues,
  getRowDimension,
  getTable,
  groupColumns,
  groupRows,
  hideColumn,
  hideColumns,
  hideRow,
  hideRows,
  isMergedCell,
  isWorksheetEmpty,
  iterCells,
  iterRows,
  iterValues,
  listComments,
  listDataValidations,
  listHyperlinks,
  listTables,
  makeWorksheet,
  mergeCells,
  moveRange,
  removeAllComments,
  removeAllConditionalFormatting,
  removeAllDataValidations,
  removeAllHyperlinks,
  removeAllMergedRanges,
  removeAllTables,
  removeCellWatches,
  removeDataValidations,
  removeHyperlink,
  removeIgnoredErrors,
  removeTable,
  setAutoFilter,
  setCell,
  setCellByCoord,
  setColumnDimension,
  setColumnWidth,
  setColumnWidths,
  setComment,
  setDefaultColumnWidth,
  setDefaultRowHeight,
  setFreezePanes,
  setHyperlink,
  setRangeValues,
  setRowDimension,
  setRowHeight,
  setRowHeights,
  setSheetTabColor,
  setSheetViewMode,
  setSheetZoom,
  ungroupColumns,
  ungroupRows,
  unhideColumn,
  unhideColumns,
  unhideRow,
  unhideRows,
  unmergeCells,
  unmergeCellsAt,
  writeRange,
} from './worksheet';
export type { MultiCellRange } from './cell-range';
export {
  expandRangeStr,
  intersectionRange,
  isCellInRange,
  isRangeInRange,
  rangeArea,
  rangeContainsCell,
  rangeContainsRange,
  rangesOverlap,
  shiftRange,
  unionRange,
} from './cell-range';
export type { ColumnDimension, RowDimension } from './dimensions';
export { makeColumnDimension, makeRowDimension } from './dimensions';
export type {
  Pane,
  PaneState,
  PaneType,
  Selection,
  SheetView,
  SheetViewMode,
} from './views';
export { freezePaneRef, makeFreezePane, makeSheetView } from './views';
export type { LegacyComment } from './comments';
export { makeLegacyComment } from './comments';
export type {
  DataValidation,
  DataValidationErrorStyle,
  DataValidationOperator,
  DataValidationType,
  ValidationCommon,
} from './data-validations';
export { makeDataValidation } from './data-validations';
export type {
  CellIsOperator,
  Cfvo,
  CfvoType,
  ConditionalFormatting,
  ConditionalFormattingRule,
  ConditionalFormattingRuleType,
  IconSetStyle,
  TextOperator,
  TimePeriod,
} from './conditional-formatting';
export { makeCfRule, makeConditionalFormatting } from './conditional-formatting';
export type { Hyperlink } from './hyperlinks';
export { makeHyperlink } from './hyperlinks';
export type { AutoFilter, FilterColumn } from './auto-filter';
export { makeAutoFilter, makeFilterColumn } from './auto-filter';
export type { TableColumn, TableDefinition, TableStyleInfo } from './table';
export { addExcelTable, makeTableColumn, makeTableDefinition } from './table';
export type { CellWatch, IgnoredError } from './errors';
export { makeCellWatch, makeIgnoredError } from './errors';
export type { OutlineProperties, PageSetupProperties, SheetProperties } from './properties';
export { makeSheetProperties } from './properties';
export type { SheetProtection } from './protection';
export { makeSheetProtection } from './protection';
export type { ProtectedRange } from './protected-ranges';
export { makeProtectedRange } from './protected-ranges';
export type {
  SortBy,
  SortCondition,
  SortIconSet,
  SortMethod,
  SortState,
} from './sort-state';
export { makeSortCondition, makeSortState } from './sort-state';
export type {
  CellSmartTag,
  CellSmartTagProperty,
  CellSmartTags,
} from './smart-tags';
export {
  makeCellSmartTag,
  makeCellSmartTagProperty,
  makeCellSmartTags,
} from './smart-tags';
export type {
  FormControl,
  OleDvAspect,
  OleObject,
  OleUpdateMode,
} from './ole-objects';
export { makeFormControl, makeOleObject } from './ole-objects';
export type { CustomSheetView, CustomSheetViewState } from './custom-sheet-views';
export { makeCustomSheetView } from './custom-sheet-views';
export type {
  CellCommentMode,
  HeaderFooter,
  HeaderFooterSection,
  PageBreak,
  PageMargins,
  PageOrder,
  PageOrientation,
  PageSetup,
  PrintErrorMode,
  PrintOptions,
} from './page-setup';
export {
  buildHeaderFooterText,
  HEADER_FOOTER_CODES,
  makeHeaderFooter,
  makePageBreak,
  makePageMargins,
  makePageSetup,
  makePrintOptions,
} from './page-setup';
export type { WebPublishItem, WorksheetCustomProperty } from './web-publish';
export { makeWebPublishItem, makeWorksheetCustomProperty } from './web-publish';
export type {
  PhoneticAlignment,
  PhoneticType,
  WorksheetPhoneticProperties,
} from './phonetic';
export { makeWorksheetPhoneticProperties } from './phonetic';
export type {
  DataConsolidate,
  DataConsolidateFunction,
  DataReference,
} from './data-consolidate';
export { makeDataConsolidate } from './data-consolidate';
export type { Scenario, ScenarioInputCell, ScenarioList } from './scenarios';
export { makeScenario, makeScenarioInputCell, makeScenarioList } from './scenarios';
