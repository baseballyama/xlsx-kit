// Public entry for openxml-js.
//
// Top-level entry: environment-neutral surface (IO / ZIP / XML / schema
// / packaging / utils foundations + the Workbook / Worksheet / Cell
// data-model + load/save). Node-only and browser-only helpers live
// behind `openxml-js/node` and `openxml-js/streaming` so this entry
// stays clean of `node:fs` and `Blob`-only imports.

// I/O abstractions (env-neutral types only — concrete helpers live in
// openxml-js/io/node and openxml-js/io/browser).
export type { BufferedSinkWriter, XlsxSink, XlsxSource } from './io';
// Packaging layer.
export type {
  CoreProperties,
  CustomProperties,
  CustomProperty,
  DefaultEntry,
  ExtendedProperties,
  Manifest,
  OverrideEntry,
  Relationship,
  Relationships,
} from './packaging';
export {
  addDefault,
  addOverride,
  appendCustomProperty,
  appendRel,
  corePropsFromBytes,
  corePropsToBytes,
  customPropsFromBytes,
  customPropsToBytes,
  extendedPropsFromBytes,
  extendedPropsToBytes,
  findAllByType,
  findById,
  findByType,
  findCustomPropertyByName,
  findOverride,
  findOverrideByContentType,
  getCustomPropertyValue,
  listCustomProperties,
  makeAsciiStringValue,
  makeBoolValue,
  makeCoreProperties,
  makeCustomProperties,
  makeDateValue,
  makeDoubleValue,
  makeExtendedProperties,
  makeFiletimeValue,
  makeIntValue,
  makeManifest,
  makeRelationships,
  makeStringValue,
  manifestFromBytes,
  manifestToBytes,
  readBoolValue,
  readDoubleValue,
  readFiletimeValue,
  readIntValue,
  readStringValue,
  relsFromBytes,
  relsToBytes,
  removeCustomProperty,
  setCustomBoolProperty,
  setCustomDateProperty,
  setCustomNumberProperty,
  setCustomStringProperty,
  setWorkbookAppVersion,
  setWorkbookApplication,
  setWorkbookCategory,
  setWorkbookCompany,
  setWorkbookCreator,
  setWorkbookDescription,
  setWorkbookHyperlinkBase,
  setWorkbookKeywords,
  setWorkbookLastModifiedBy,
  setWorkbookManager,
  setWorkbookSubject,
  setWorkbookTitle,
} from './packaging';
// Phase 3 (read / write). Currently the loadWorkbook minimum skeleton —
// reads manifest + workbook.xml + sheet rels and produces a Workbook
// shell. Cell content is filled in by later iterations.
export type { LoadOptions } from './public/load';
export { loadWorkbook } from './public/load';
export type { SaveOptions } from './public/save';
export { saveWorkbook, workbookToBytes } from './public/save';
// Schema layer.
export type { AttrDef, ElementDef, Primitive, Schema } from './schema';
export { defineSchema, fromTree, toTree } from './schema';
// Utility surfaces — coordinate / datetime / units / inference / escape /
// exception types.
export type { CellCoordinate, CellCoordinateNumeric, CellRangeBoundaries } from './utils/coordinate';
export {
  cellRangeFromCells,
  expandRangeStr,
  intersectionRangeStr,
  isCellInRange,
  isRangeInRange,
  rangeAreaStr,
  rangeDimensionsStr,
  rangesOverlapStr,
  shiftRangeStr,
  unionRangeStr,
} from './worksheet/cell-range';
export {
  boundariesToRangeString,
  columnIndexFromLetter,
  columnLetterFromIndex,
  coordinateFromString,
  coordinateToTuple,
  formatSheetQualifiedRef,
  isValidCellRef,
  isValidColumnLetter,
  isValidColumnNumber,
  isValidRangeRef,
  isValidRowNumber,
  MAX_COL,
  MAX_ROW,
  parseSheetRange,
  rangeBoundaries,
  tupleToCoordinate,
} from './utils/coordinate';
export type { ExcelEpoch } from './utils/datetime';
export {
  dateToExcel,
  durationToExcel,
  excelToDate,
  excelToDuration,
  fromIso8601,
  MAC_EPOCH_MS,
  toIso8601,
  WINDOWS_EPOCH_MS,
} from './utils/datetime';
export { cssRecordToInlineStyle } from './utils/css';
export { escapeCellString, unescapeCellString } from './utils/escape';
export {
  OpenXmlError,
  OpenXmlInvalidWorkbookError,
  OpenXmlIoError,
  OpenXmlNotImplementedError,
  OpenXmlSchemaError,
} from './utils/exceptions';
export type { CellDataType } from './utils/inference';
export { ERROR_CODES, inferCellType } from './utils/inference';
export {
  cmFromEmu,
  EMU_PER_CM,
  EMU_PER_INCH,
  EMU_PER_PIXEL,
  EMU_PER_POINT,
  emuFromCm,
  emuFromInch,
  emuFromPoint,
  emuFromPx,
  inchFromEmu,
  pixelToPoint,
  pointFromEmu,
  pointToPixel,
  pxFromEmu,
} from './utils/units';
// XML layer.
export type {
  SaxEvent,
  SaxInput,
  SerializeOptions,
  XmlNode,
  XmlStreamWriter,
  XmlStreamWriterOptions,
} from './xml';
export {
  appendChild,
  createXmlStreamWriter,
  el,
  elNs,
  findChild,
  findChildren,
  iterParse,
  parseQName,
  parseXml,
  qname,
  serializeXml,
} from './xml';
// ZIP layer.
export type { ZipArchive, ZipWriter, StreamingEntryWriter } from './zip';
export { createZipWriter, openZip } from './zip';

// Cell value-model and helpers.
export type {
  Cell,
  CellValue,
  DataTableFormulaOpts,
  ExcelErrorCode,
  FormulaKind,
  FormulaValue,
  MergedCell,
} from './cell/cell';
export {
  bindValue,
  cellValueAsBoolean,
  cellValueAsDate,
  cellValueAsNumber,
  cellValueAsString,
  getCachedFormulaValue,
  getCoordinate,
  getFormulaText,
  isDurationValue,
  isEmptyCell,
  isErrorValue,
  isFormulaCell,
  isFormulaValue,
  isRichTextCell,
  isRichTextValue,
  makeCell,
  makeDurationValue,
  makeErrorValue,
  setArrayFormula,
  setCellValue,
  setDataTableFormula,
  setFormula,
  setSharedFormula,
} from './cell/cell';
// Inline rich-text helpers — composed inside CellValue when the cell
// carries multi-format text (vs. plain string → sharedStrings).
export type { InlineFont, RichText, TextRun } from './cell/rich-text';
export { makeRichText, makeTextRun, richTextToString } from './cell/rich-text';

// Worksheet model + helpers (mergeCells / freezePanes / dimensions /
// hyperlinks / data-validations / autoFilter / tables / comments /
// conditional-formatting are reachable through the Worksheet object).
export type { CellsByKindCounts, Worksheet } from './worksheet/worksheet';
export type { CellSummary, WorkbookOverview, WorkbookSheetOverview } from './workbook/workbook';
export {
  addCellWatch,
  addConditionalFormatting,
  addDataValidation,
  addIgnoredError,
  addTable,
  appendRow,
  appendRows,
  countCells,
  countCellsByKind,
  deleteCell,
  editCommentAuthor,
  editCommentText,
  findCells,
  findCommentsByAuthor,
  findFirstCell,
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
  hideColumn,
  hideColumns,
  hideRow,
  hideRows,
  isMergedCell,
  isWorksheetEmpty,
  iterCells as iterWorksheetCells,
  iterRows as iterWorksheetRows,
  iterValues as iterWorksheetValues,
  listComments,
  listDataValidations,
  listHyperlinks,
  listTables,
  makeWorksheet,
  mergeCells,
  removeAllComments,
  removeAllConditionalFormatting,
  removeAllDataValidations,
  removeAllHyperlinks,
  removeAllMergedRanges,
  removeAllTables,
  removeCellWatches,
  removeComment,
  removeDataValidations,
  removeHyperlink,
  removeIgnoredErrors,
  removeSheetTabColor,
  removeTable,
  renameCommentAuthor,
  readRangeAsObjects,
  writeRangeFromObjects,
  replaceCellValues,
  replaceInRange,
  applyToRange,
  autofitColumn,
  autofitColumns,
  clearAllCells,
  clearRange,
  copyRange,
  moveRange,
  collapseColumnGroup,
  collapseRowGroup,
  expandColumnGroup,
  expandRowGroup,
  freezeColumns,
  freezeFirstColumn,
  freezeFirstRow,
  freezeFirstRowAndColumn,
  freezePanes,
  freezeRows,
  groupColumns,
  groupRows,
  setActiveCell,
  setAutoFilter,
  setCell,
  setCellArrayFormula,
  setCellByCoord,
  setCellFormula,
  setCellRichText,
  setSelectedRange,
  setColumnDimension,
  setColumnWidth,
  setColumnWidths,
  setDefaultColumnWidth,
  setDefaultRowHeight,
  setComment,
  setFreezePanes,
  setHyperlink,
  setRangeValues,
  setRightToLeft,
  setRowDimension,
  setRowHeight,
  setRowHeights,
  setSheetTabColor,
  setSheetViewMode,
  setSheetZoom,
  setShowFormulas,
  setShowGridLines,
  setShowRowColHeaders,
  setShowZeros,
  unfreezePanes,
  ungroupColumns,
  ungroupRows,
  unhideColumn,
  unhideColumns,
  unhideRow,
  unhideRows,
  unmergeCells,
  unmergeCellsAt,
  writeRange,
} from './worksheet/worksheet';
export type { ColumnDimension, RowDimension } from './worksheet/dimensions';
export { makeColumnDimension, makeRowDimension } from './worksheet/dimensions';
export type {
  DataValidation,
  DataValidationErrorStyle,
  DataValidationOperator,
  DataValidationType,
  ValidationCommon,
} from './worksheet/data-validations';
export {
  addCustomValidation,
  addDateValidation,
  addListValidation,
  addNumberValidation,
  makeDataValidation,
} from './worksheet/data-validations';
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
} from './worksheet/conditional-formatting';
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
} from './worksheet/conditional-formatting';
export type { Hyperlink } from './worksheet/hyperlinks';
export {
  addInternalHyperlink,
  addMailtoHyperlink,
  addUrlHyperlink,
  makeHyperlink,
} from './worksheet/hyperlinks';
export type {
  ChartReference,
  Drawing,
  DrawingItem,
  PictureReference,
} from './drawing/drawing';
export {
  addChartAt,
  addImageAt,
  listChartsOnSheet,
  listImagesOnSheet,
  makeChartDrawingItem,
  makeDrawing,
  makePictureDrawingItem,
  removeAllCharts,
  removeAllDrawingItems,
  removeAllImages,
} from './drawing/drawing';
export type { XlsxImage } from './drawing/image';
export { loadImage } from './drawing/image';
export type { DrawingAnchor } from './drawing/anchor';
export { makeOneCellAnchor } from './drawing/anchor';
export type { AutoFilter, FilterColumn } from './worksheet/auto-filter';
export {
  addAutoFilter,
  addAutoFilterColumn,
  makeAutoFilter,
  makeFilterColumn,
  removeAutoFilter,
} from './worksheet/auto-filter';
export type { TableColumn, TableDefinition, TableStyleInfo } from './worksheet/table';
export { addExcelTable, addTableFromObjects, makeTableColumn, makeTableDefinition } from './worksheet/table';
export { getRangeAsCsv, getWorksheetAsCsv, parseCsv, parseCsvToRange } from './worksheet/csv';
export type { CellWatch, IgnoredError } from './worksheet/errors';
export { makeCellWatch, makeIgnoredError } from './worksheet/errors';
export type { OutlineProperties, PageSetupProperties, SheetProperties } from './worksheet/properties';
export { makeSheetProperties } from './worksheet/properties';
export type { SheetProtection } from './worksheet/protection';
export {
  isSheetProtected,
  makeSheetProtection,
  protectSheet,
  unprotectSheet,
} from './worksheet/protection';
export type { ProtectedRange } from './worksheet/protected-ranges';
export { makeProtectedRange } from './worksheet/protected-ranges';
export type {
  SortBy,
  SortCondition,
  SortIconSet,
  SortMethod,
  SortState,
} from './worksheet/sort-state';
export { makeSortCondition, makeSortState } from './worksheet/sort-state';
export type {
  CellSmartTag,
  CellSmartTagProperty,
  CellSmartTags,
} from './worksheet/smart-tags';
export {
  makeCellSmartTag,
  makeCellSmartTagProperty,
  makeCellSmartTags,
} from './worksheet/smart-tags';
export type {
  FormControl,
  OleDvAspect,
  OleObject,
  OleUpdateMode,
} from './worksheet/ole-objects';
export { makeFormControl, makeOleObject } from './worksheet/ole-objects';
export type { CustomSheetView, CustomSheetViewState } from './worksheet/custom-sheet-views';
export { makeCustomSheetView } from './worksheet/custom-sheet-views';
export type {
  CellCommentMode,
  HeaderFooter,
  PageBreak,
  PageMargins,
  PageOrder,
  PageOrientation,
  PageSetup,
  PrintErrorMode,
  PrintOptions,
} from './worksheet/page-setup';
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
  setPrintCentered,
  setPrintGridLines,
  setPrintHeadings,
  setPageMargins,
  setPageOrientation,
  setPaperSize,
  setPrintScale,
} from './worksheet/page-setup';
export type { HeaderFooterSection } from './worksheet/page-setup';
export type { WebPublishItem, WorksheetCustomProperty } from './worksheet/web-publish';
export { makeWebPublishItem, makeWorksheetCustomProperty } from './worksheet/web-publish';
export type {
  PhoneticAlignment,
  PhoneticType,
  WorksheetPhoneticProperties,
} from './worksheet/phonetic';
export { makeWorksheetPhoneticProperties } from './worksheet/phonetic';
export type {
  DataConsolidate,
  DataConsolidateFunction,
  DataReference,
} from './worksheet/data-consolidate';
export { makeDataConsolidate } from './worksheet/data-consolidate';
export type { Scenario, ScenarioInputCell, ScenarioList } from './worksheet/scenarios';
export { makeScenario, makeScenarioInputCell, makeScenarioList } from './worksheet/scenarios';

// Style value objects (Color / Font / Fill / Border / Alignment /
// Protection / NumberFormat) + the cell ↔ stylesheet bridge.
export type {
  Alignment,
  Border,
  Color,
  Fill,
  Font,
  GradientFill,
  GradientStop,
  PatternFill,
  Protection,
  Side,
} from './styles';
export type { DifferentialStyle, NamedStyle, StylesheetNamedStyle } from './styles';
export {
  addBorder,
  addCellStyleXf,
  addCellXf,
  addDxf,
  addFill,
  addFont,
  addNamedStyle,
  addNumFmt,
  alignCellHorizontal,
  alignCellVertical,
  alignmentToCss,
  applyBuiltinStyle,
  applyNamedStyle,
  borderToCss,
  BUILTIN_NAMED_STYLES,
  cellStyleToCss,
  centerCell,
  clearCellBackground,
  clearCellStyle,
  clearRangeStyle,
  cloneCellStyle,
  copyCellStyle,
  defaultCellXf,
  ensureBuiltinStyle,
  fillToCss,
  fontToCss,
  formatAsHeader,
  indentCell,
  listBorders,
  listCellStyleXfs,
  listCellXfs,
  listFills,
  listFonts,
  rotateCellText,
  setBold,
  setCellBackgroundColor,
  setFontColor,
  setFontName,
  setFontSize,
  setItalic,
  setStrikethrough,
  setUnderline,
  wrapCellText,
  getCellAlignment,
  getCellBorder,
  getCellFill,
  getCellFont,
  getCellNumberFormat,
  getCellProtection,
  adjustLightness,
  adjustSaturation,
  colorToHex,
  contrastRatio,
  darken,
  hexToHsl,
  hslToHex,
  lighten,
  luminance,
  makeAlignment,
  makeBorder,
  makeColor,
  makeDifferentialStyle,
  makeFill,
  makeFont,
  makeGradientFill,
  makeGradientStop,
  makePatternFill,
  mixColors,
  makeProtection,
  makeSide,
  makeStylesheet,
  pickReadableTextColor,
  rgbColor,
  rotateHue,
  setCellAlignment,
  setCellAsCurrency,
  setCellAsDate,
  setCellAsNumber,
  setCellAsPercent,
  setCellBorder,
  setCellBorderAll,
  setCellFill,
  setCellFont,
  setCellNumberFormat,
  setCellProtection,
  setCellStyle,
  setRangeBackgroundColor,
  setRangeBorderBox,
  setRangeFont,
  setRangeNumberFormat,
  setRangeAlignment,
  setRangeProtection,
  setRangeStyle,
  setRangeWrapText,
} from './styles';

// Workbook root model.
export type { SheetRef, SheetState, Workbook, WorkbookStats } from './workbook/workbook';
export {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  createWorkbookFromCsv,
  createWorkbookFromCsvBundle,
  createWorkbookFromObjects,
  duplicateSheet,
  getActiveSheet,
  findCellInWorkbook,
  findCellsInWorkbook,
  findTable,
  getAllCharts,
  getAllComments,
  getAllConditionalFormatting,
  getAllDataValidations,
  getAllHyperlinks,
  getAllImages,
  getAllMergedRanges,
  getAllTables,
  getChartsheet,
  getSheet,
  describeWorkbook,
  getCellAtAddress,
  getCellSummary,
  getRangeValuesAtAddress,
  getValueAtAddress,
  isWorkbookEmpty,
  setCellAtAddress,
  setRangeValuesAtAddress,
  getSheetByIndex,
  getSheetState,
  getWorkbookAsCsvBundle,
  getWorkbookAsCsvRecord,
  getWorkbookCellsByKind,
  getWorkbookStats,
  hideSheet,
  isValidSheetTitle,
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
  setSheetState,
  setSheetStates,
  sheetNames,
  showAllSheets,
  showSheet,
  swapSheets,
  validateSheetTitle,
  veryHideSheet,
} from './workbook/workbook';
export type { DefinedName, DefinedNameTarget } from './workbook/defined-names';
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
} from './workbook/defined-names';
export type { WorkbookProtection } from './workbook/protection';
export {
  isWorkbookProtected,
  makeWorkbookProtection,
  protectWorkbook,
  unprotectWorkbook,
} from './workbook/protection';
export type { SharedStringsTable } from './workbook/shared-strings';
export {
  addSharedString,
  getSharedStringAt,
  getSharedStringIndex,
  makeSharedStrings,
  sharedStringCount,
} from './workbook/shared-strings';
export type {
  CustomViewShowComments,
  CustomViewShowObjects,
  CustomWorkbookView,
  WorkbookView,
  WorkbookViewVisibility,
} from './workbook/views';
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
} from './workbook/views';
export type { CalcMode, CalcProperties, RefMode } from './workbook/calc-properties';
export {
  makeCalcProperties,
  setCalcMode,
  setCalcOnSave,
  setFullCalcOnLoad,
  setFullPrecision,
  setIterativeCalc,
} from './workbook/calc-properties';
export type {
  ShowObjectsMode,
  UpdateLinksMode,
  WorkbookProperties,
} from './workbook/workbook-properties';
export {
  makeWorkbookProperties,
  setDate1904,
  setFilterPrivacy,
  setUpdateLinksMode,
  setWorkbookCodeName,
} from './workbook/workbook-properties';
export type { FileVersion } from './workbook/file-version';
export { makeFileVersion } from './workbook/file-version';
export type { FileSharing } from './workbook/file-sharing';
export { makeFileSharing } from './workbook/file-sharing';
export type { FileRecoveryProperties } from './workbook/file-recovery';
export { makeFileRecoveryProperties } from './workbook/file-recovery';
export type { SmartTagProperties, SmartTagShowMode, SmartTagType } from './workbook/smart-tags';
export { makeSmartTagProperties, makeSmartTagType } from './workbook/smart-tags';
export type { FunctionGroup, FunctionGroups } from './workbook/function-groups';
export { makeFunctionGroup, makeFunctionGroups } from './workbook/function-groups';
