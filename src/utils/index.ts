// Utility surfaces — coordinate / datetime / units / inference / escape /
// css / exception types.

export type { CellCoordinate, CellCoordinateNumeric, CellRangeBoundaries } from './coordinate';
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
} from './coordinate';
export type { ExcelEpoch } from './datetime';
export {
  dateToExcel,
  durationToExcel,
  excelToDate,
  excelToDuration,
  fromIso8601,
  MAC_EPOCH_MS,
  toIso8601,
  WINDOWS_EPOCH_MS,
} from './datetime';
export { cssRecordToInlineStyle } from './css';
export { escapeCellString, unescapeCellString } from './escape';
export type { OpenXmlErrorOptions } from './exceptions';
export {
  OpenXmlDecompressionBombError,
  OpenXmlError,
  OpenXmlInvalidWorkbookError,
  OpenXmlIoError,
  OpenXmlNotImplementedError,
  OpenXmlSchemaError,
} from './exceptions';
export type { CellDataType } from './inference';
export { ERROR_CODES, inferCellType } from './inference';
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
} from './units';
