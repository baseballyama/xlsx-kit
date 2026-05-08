// Public surface for drawings (charts/images embedded in a worksheet),
// anchors, image bytes, and DML shape properties.

export type {
  ChartReference,
  Drawing,
  DrawingItem,
  PictureReference,
} from './drawing';
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
} from './drawing';
export type { XlsxImage, XlsxImageFormat } from './image';
export { loadImage } from './image';
export type {
  AnchorMarker,
  DrawingAnchor,
  Point2D,
  PositiveSize2D,
} from './anchor';
export { makeOneCellAnchor } from './anchor';
export type { BlackWhiteMode, ShapeProperties, Transform2D } from './dml/shape-properties';
export { makeShapeProperties } from './dml/shape-properties';
