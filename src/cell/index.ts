// Public surface for the Cell value-model (CellValue + formula kinds +
// inline rich-text).

export type {
  Cell,
  CellValue,
  DataTableFormulaOpts,
  ExcelErrorCode,
  FormulaKind,
  FormulaValue,
  MergedCell,
} from './cell';
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
} from './cell';
export type {
  InlineFont,
  InlineUnderline,
  InlineVertAlign,
  RichText,
  TextRun,
} from './rich-text';
export { makeRichText, makeTextRun, richTextToString } from './rich-text';
