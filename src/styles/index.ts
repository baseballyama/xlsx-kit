// Public surface for the style value-objects + cell ↔ stylesheet bridge.
// Per docs/plan/04-core-model.md §3 — Color / Font / Fill / Border /
// Alignment / Protection / NumberFormat are plain objects with `make*`
// factories, and Stylesheet pools dedup equal values via stable keys.

export type { Alignment, HorizontalAlignment, VerticalAlignment } from './alignment';
export { makeAlignment } from './alignment';
export type { Border, Side, SideStyle } from './borders';
export { makeBorder, makeSide } from './borders';
export {
  alignCellHorizontal,
  alignCellVertical,
  applyBuiltinStyle,
  applyNamedStyle,
  centerCell,
  clearCellBackground,
  cloneCellStyle,
  copyCellStyle,
  formatAsHeader,
  indentCell,
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
  setRangeStyle,
} from './cell-style';
export type { DifferentialStyle } from './differential';
export { addDxf, makeDifferentialStyle } from './differential';
export type { NamedStyle, StylesheetNamedStyle } from './named-styles';
export { addNamedStyle, BUILTIN_NAMED_STYLES, ensureBuiltinStyle } from './named-styles';
export type { Color } from './colors';
export {
  adjustLightness,
  adjustSaturation,
  colorToHex,
  contrastRatio,
  darken,
  hexToHsl,
  hslToHex,
  lighten,
  luminance,
  makeColor,
  mixColors,
  normaliseRgb,
  pickReadableTextColor,
  resolveIndexedColor,
  rgbColor,
  rotateHue,
} from './colors';
export type { Fill, GradientFill, GradientFillType, GradientStop, PatternFill, PatternType } from './fills';
export { makeFill, makeGradientFill, makeGradientStop, makePatternFill } from './fills';
export type { Font, FontScheme, UnderlineStyle, VertAlign } from './fonts';
export { DEFAULT_FONT, fontToCss, makeFont } from './fonts';
export type { Protection } from './protection';
export { makeProtection } from './protection';
export type { CellXf, Stylesheet } from './stylesheet';
export {
  addBorder,
  addCellStyleXf,
  addCellXf,
  addFill,
  addFont,
  addNumFmt,
  defaultCellXf,
  listBorders,
  listCellStyleXfs,
  listCellXfs,
  listFills,
  listFonts,
  makeStylesheet,
} from './stylesheet';
