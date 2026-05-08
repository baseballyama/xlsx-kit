// Public surface for the Worksheet model + cell-range / dimensions / views /
// comments / hyperlinks / data-validations / conditional-formatting /
// auto-filter / tables / page-setup / protection / errors / smart-tags /
// ole-objects / sort-state / scenarios / data-consolidate / web-publish /
// phonetic / protected-ranges / properties / custom-sheet-views / csv / html
// / markdown / text exporters.

export type {
  CellsByKindCounts,
  ColumnAggregates,
  IterRowsOptions,
  PivotAggregate,
  Worksheet,
} from './worksheet';
export {
  addCellWatch,
  addColumn,
  addConditionalFormatting,
  addDataValidation,
  addIgnoredError,
  addTable,
  appendRow,
  appendRows,
  applyToRange,
  autofitColumn,
  autofitColumns,
  clearAllCells,
  clearRange,
  collapseColumnGroup,
  collapseRowGroup,
  columnAggregates,
  columnIndexOf,
  copyRange,
  countCells,
  countCellsByKind,
  countRows,
  deleteCell,
  editCommentAuthor,
  editCommentText,
  everyRow,
  expandColumnGroup,
  expandRowGroup,
  fillColumn,
  filterRange,
  findCells,
  findCommentsByAuthor,
  findFirstCell,
  findRow,
  forEachRow,
  freezeColumns,
  freezeFirstColumn,
  freezeFirstRow,
  freezeFirstRowAndColumn,
  freezePanes,
  freezeRows,
  getAutoFilter,
  getCell,
  getCellAddress,
  getCellByCoord,
  getCellComment,
  getCellHyperlink,
  getCellsInColumn,
  getCellsInRange,
  getCellsInRow,
  getColumnDimension,
  getColumnValues,
  getComment,
  getConditionalFormatting,
  getDataExtent,
  getDataExtentRef,
  getDistinctValuesInColumn,
  getDistinctValuesInRow,
  getFreezePanes,
  getHeaders,
  getHyperlink,
  getMaxCol,
  getMaxRow,
  getMergedCells,
  getMergedRangeAt,
  getNonEmptyCellCount,
  getPopulatedColumnIndices,
  getPopulatedRowIndices,
  getRangeAddress,
  getRangeValues,
  getRowDimension,
  getRowValues,
  getTable,
  groupBy,
  groupColumns,
  groupRows,
  hasColumn,
  hideColumn,
  hideColumns,
  hideRow,
  hideRows,
  indexOfRow,
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
  mapRange,
  mergeCells,
  moveRange,
  pivotTable,
  pluckColumn,
  readRangeAsObjects,
  reduceRange,
  removeAllComments,
  removeAllConditionalFormatting,
  removeAllDataValidations,
  removeAllHyperlinks,
  removeAllMergedRanges,
  removeAllTables,
  removeCellWatches,
  removeColumn,
  removeComment,
  removeDataValidations,
  removeHyperlink,
  removeIgnoredErrors,
  removeSheetTabColor,
  removeTable,
  renameColumn,
  renameColumns,
  renameCommentAuthor,
  reorderColumns,
  replaceCellValues,
  replaceInRange,
  setActiveCell,
  setAutoFilter,
  setCell,
  setCellArrayFormula,
  setCellByCoord,
  setCellFormula,
  setCellRichText,
  setColumnDimension,
  setColumnWidth,
  setColumnWidths,
  setComment,
  setDefaultColumnWidth,
  setDefaultRowHeight,
  setFreezePanes,
  setHyperlink,
  setRangeValues,
  setRightToLeft,
  setRowDimension,
  setRowHeight,
  setRowHeights,
  setSelectedRange,
  setSheetTabColor,
  setSheetViewMode,
  setSheetZoom,
  setShowFormulas,
  setShowGridLines,
  setShowRowColHeaders,
  setShowZeros,
  someRow,
  sortRange,
  tabularData,
  unfreezePanes,
  ungroupColumns,
  ungroupRows,
  unhideColumn,
  unhideColumns,
  unhideRow,
  unhideRows,
  uniqueColumn,
  unmergeCells,
  unmergeCellsAt,
  writeRange,
  writeRangeFromObjects,
} from './worksheet';
export type { MultiCellRange } from './cell-range';
export {
  cellRangeFromCells,
  expandRangeStr,
  intersectionRange,
  intersectionRangeStr,
  isCellInRange,
  isRangeInRange,
  makeMultiCellRange,
  multiCellRangeToString,
  parseMultiCellRange,
  rangeArea,
  rangeAreaStr,
  rangeContainsCell,
  rangeContainsRange,
  rangeDimensionsStr,
  rangesOverlap,
  rangesOverlapStr,
  shiftRange,
  shiftRangeStr,
  unionRange,
  unionRangeStr,
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
export {
  addCustomValidation,
  addDateValidation,
  addListValidation,
  addNumberValidation,
  makeDataValidation,
} from './data-validations';
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
export {
  addAverageRule,
  addCellIsRule,
  addColorScaleRule,
  addDataBarRule,
  addDuplicateValuesRule,
  addFormulaRule,
  addIconSetRule,
  addTextRule,
  addTopNRule,
  makeCfRule,
  makeConditionalFormatting,
} from './conditional-formatting';
export type { Hyperlink } from './hyperlinks';
export {
  addInternalHyperlink,
  addMailtoHyperlink,
  addUrlHyperlink,
  makeHyperlink,
} from './hyperlinks';
export type { AutoFilter, FilterColumn } from './auto-filter';
export {
  addAutoFilter,
  addAutoFilterColumn,
  makeAutoFilter,
  makeFilterColumn,
  removeAutoFilter,
} from './auto-filter';
export type { TableColumn, TableDefinition, TableStyleInfo } from './table';
export { addExcelTable, addTableFromObjects, makeTableColumn, makeTableDefinition } from './table';
export { getRangeAsCsv, getWorksheetAsCsv, parseCsv, parseCsvToRange } from './csv';
export { getWorksheetAsHtml, worksheetToHtml } from './html';
export type { WorksheetToJsonOptions } from './json';
export { worksheetToJson } from './json';
export { getWorksheetAsMarkdownTable, worksheetToMarkdownTable } from './markdown';
export { getWorksheetAsTextTable, worksheetToTextTable } from './text';
export type { CellWatch, IgnoredError } from './errors';
export { makeCellWatch, makeIgnoredError } from './errors';
export type { OutlineProperties, PageSetupProperties, SheetProperties } from './properties';
export { makeSheetProperties } from './properties';
export type { SheetProtection } from './protection';
export {
  isSheetProtected,
  makeSheetProtection,
  protectSheet,
  unprotectSheet,
} from './protection';
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
  addColBreak,
  addRowBreak,
  buildHeaderFooterText,
  HEADER_FOOTER_CODES,
  makeHeaderFooter,
  makePageBreak,
  makePageMargins,
  makePageSetup,
  makePrintOptions,
  setFitToPage,
  setFooter,
  setFooterText,
  setHeader,
  setHeaderText,
  setPageMargins,
  setPageOrientation,
  setPaperSize,
  setPrintCentered,
  setPrintGridLines,
  setPrintHeadings,
  setPrintScale,
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
