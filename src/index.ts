// Public entry for openxml-js.
//
// Phase 1 surface: the foundation layer (IO, ZIP, XML, schema, packaging,
// utils). Higher-level Workbook / Worksheet / Cell APIs land in phase 2+
// per docs/plan/04-core-model.md.
//
// Sub-path entrypoints (openxml-js/io/node etc.) carry environment-
// specific helpers; this top-level entry only re-exports environment-
// neutral types so a single import works in both Node and the browser.

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
export type { ZipArchive, ZipWriter } from './zip';
export { createZipWriter, openZip } from './zip';
