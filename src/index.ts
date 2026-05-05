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
  boundariesToRangeString,
  columnIndexFromLetter,
  columnLetterFromIndex,
  coordinateFromString,
  coordinateToTuple,
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
  getCoordinate,
  isEmptyCell,
  isFormulaCell,
  isRichTextCell,
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
export type { Worksheet } from './worksheet/worksheet';
export {
  addCellWatch,
  addConditionalFormatting,
  addDataValidation,
  addIgnoredError,
  addTable,
  appendRow,
  countCells,
  deleteCell,
  getAutoFilter,
  getCell,
  getCellByCoord,
  getColumnDimension,
  getComment,
  getConditionalFormatting,
  getFreezePanes,
  getHyperlink,
  getMaxCol,
  getMaxRow,
  getMergedCells,
  getRowDimension,
  getTable,
  hideColumn,
  hideRow,
  isMergedCell,
  iterRows as iterWorksheetRows,
  iterValues as iterWorksheetValues,
  makeWorksheet,
  mergeCells,
  removeCellWatches,
  removeComment,
  removeDataValidations,
  removeHyperlink,
  removeIgnoredErrors,
  removeTable,
  setAutoFilter,
  setCell,
  setCellByCoord,
  setColumnDimension,
  setColumnWidth,
  setComment,
  setFreezePanes,
  setHyperlink,
  setRowDimension,
  setRowHeight,
  unmergeCells,
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
  ConditionalFormatting,
  ConditionalFormattingRule,
  ConditionalFormattingRuleType,
  TextOperator,
  TimePeriod,
} from './worksheet/conditional-formatting';
export {
  addAverageRule,
  addCellIsRule,
  addDuplicateValuesRule,
  addFormulaRule,
  addTextRule,
  addTopNRule,
  makeCfRule,
  makeConditionalFormatting,
} from './worksheet/conditional-formatting';
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
  makeHeaderFooter,
  makePageBreak,
  makePageMargins,
  makePageSetup,
  makePrintOptions,
  setFitToPage,
  setFooter,
  setHeader,
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
  BUILTIN_NAMED_STYLES,
  defaultCellXf,
  ensureBuiltinStyle,
  getCellAlignment,
  getCellBorder,
  getCellFill,
  getCellFont,
  getCellNumberFormat,
  getCellProtection,
  makeAlignment,
  makeBorder,
  makeColor,
  makeDifferentialStyle,
  makeFill,
  makeFont,
  makeGradientFill,
  makeGradientStop,
  makePatternFill,
  makeProtection,
  makeSide,
  makeStylesheet,
  rgbColor,
  setCellAlignment,
  setCellBorder,
  setCellFill,
  setCellFont,
  setCellNumberFormat,
  setCellProtection,
} from './styles';

// Workbook root model.
export type { SheetRef, SheetState, Workbook } from './workbook/workbook';
export {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getActiveSheet,
  getChartsheet,
  getSheet,
  getSheetByIndex,
  jsonReplacer,
  jsonReviver,
  listCustomXmlParts,
  removeSheet,
  setActiveSheet,
  sheetNames,
} from './workbook/workbook';
export type { WorkbookProtection } from './workbook/protection';
export {
  isWorkbookProtected,
  makeWorkbookProtection,
  protectWorkbook,
  unprotectWorkbook,
} from './workbook/protection';
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
  setShowSheetTabs,
  setTabRatio,
  setWorkbookWindow,
} from './workbook/views';
export type { CalcMode, CalcProperties, RefMode } from './workbook/calc-properties';
export { makeCalcProperties } from './workbook/calc-properties';
export type {
  ShowObjectsMode,
  UpdateLinksMode,
  WorkbookProperties,
} from './workbook/workbook-properties';
export { makeWorkbookProperties } from './workbook/workbook-properties';
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
