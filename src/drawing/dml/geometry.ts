// DrawingML geometry. Per docs/plan/08-charts-drawings.md §4.5.
//
// Geometry has two kinds: a preset shape pulled from the ECMA-376
// gallery (one of the 187 `ST_PresetShapeType` names below) and a
// custom path expressed as a sequence of move/line/arc/bezier/close
// commands. openpyxl drops the path command list on read; the model
// below preserves it so the "custGeom round-trip" acceptance criterion
// can pass.

import type { Point2D } from './shape-properties';

/**
 * ECMA-376 §20.1.10.55 preset shape catalogue (`ST_PresetShapeType`,
 * 187 entries). Pulled from openpyxl's `PresetGeometry2D.prst` Set so
 * the test corpus stays directly comparable.
 */
export const PRESET_SHAPE_NAMES: ReadonlyArray<string> = [
  'line',
  'lineInv',
  'triangle',
  'rtTriangle',
  'rect',
  'diamond',
  'parallelogram',
  'trapezoid',
  'nonIsoscelesTrapezoid',
  'pentagon',
  'hexagon',
  'heptagon',
  'octagon',
  'decagon',
  'dodecagon',
  'star4',
  'star5',
  'star6',
  'star7',
  'star8',
  'star10',
  'star12',
  'star16',
  'star24',
  'star32',
  'roundRect',
  'round1Rect',
  'round2SameRect',
  'round2DiagRect',
  'snipRoundRect',
  'snip1Rect',
  'snip2SameRect',
  'snip2DiagRect',
  'plaque',
  'ellipse',
  'teardrop',
  'homePlate',
  'chevron',
  'pieWedge',
  'pie',
  'blockArc',
  'donut',
  'noSmoking',
  'rightArrow',
  'leftArrow',
  'upArrow',
  'downArrow',
  'stripedRightArrow',
  'notchedRightArrow',
  'bentUpArrow',
  'leftRightArrow',
  'upDownArrow',
  'leftUpArrow',
  'leftRightUpArrow',
  'quadArrow',
  'leftArrowCallout',
  'rightArrowCallout',
  'upArrowCallout',
  'downArrowCallout',
  'leftRightArrowCallout',
  'upDownArrowCallout',
  'quadArrowCallout',
  'bentArrow',
  'uturnArrow',
  'circularArrow',
  'leftCircularArrow',
  'leftRightCircularArrow',
  'curvedRightArrow',
  'curvedLeftArrow',
  'curvedUpArrow',
  'curvedDownArrow',
  'swooshArrow',
  'cube',
  'can',
  'lightningBolt',
  'heart',
  'sun',
  'moon',
  'smileyFace',
  'irregularSeal1',
  'irregularSeal2',
  'foldedCorner',
  'bevel',
  'frame',
  'halfFrame',
  'corner',
  'diagStripe',
  'chord',
  'arc',
  'leftBracket',
  'rightBracket',
  'leftBrace',
  'rightBrace',
  'bracketPair',
  'bracePair',
  'straightConnector1',
  'bentConnector2',
  'bentConnector3',
  'bentConnector4',
  'bentConnector5',
  'curvedConnector2',
  'curvedConnector3',
  'curvedConnector4',
  'curvedConnector5',
  'callout1',
  'callout2',
  'callout3',
  'accentCallout1',
  'accentCallout2',
  'accentCallout3',
  'borderCallout1',
  'borderCallout2',
  'borderCallout3',
  'accentBorderCallout1',
  'accentBorderCallout2',
  'accentBorderCallout3',
  'wedgeRectCallout',
  'wedgeRoundRectCallout',
  'wedgeEllipseCallout',
  'cloudCallout',
  'cloud',
  'ribbon',
  'ribbon2',
  'ellipseRibbon',
  'ellipseRibbon2',
  'leftRightRibbon',
  'verticalScroll',
  'horizontalScroll',
  'wave',
  'doubleWave',
  'plus',
  'flowChartProcess',
  'flowChartDecision',
  'flowChartInputOutput',
  'flowChartPredefinedProcess',
  'flowChartInternalStorage',
  'flowChartDocument',
  'flowChartMultidocument',
  'flowChartTerminator',
  'flowChartPreparation',
  'flowChartManualInput',
  'flowChartManualOperation',
  'flowChartConnector',
  'flowChartPunchedCard',
  'flowChartPunchedTape',
  'flowChartSummingJunction',
  'flowChartOr',
  'flowChartCollate',
  'flowChartSort',
  'flowChartExtract',
  'flowChartMerge',
  'flowChartOfflineStorage',
  'flowChartOnlineStorage',
  'flowChartMagneticTape',
  'flowChartMagneticDisk',
  'flowChartMagneticDrum',
  'flowChartDisplay',
  'flowChartDelay',
  'flowChartAlternateProcess',
  'flowChartOffpageConnector',
  'actionButtonBlank',
  'actionButtonHome',
  'actionButtonHelp',
  'actionButtonInformation',
  'actionButtonForwardNext',
  'actionButtonBackPrevious',
  'actionButtonEnd',
  'actionButtonBeginning',
  'actionButtonReturn',
  'actionButtonDocument',
  'actionButtonSound',
  'actionButtonMovie',
  'gear6',
  'gear9',
  'funnel',
  'mathPlus',
  'mathMinus',
  'mathMultiply',
  'mathDivide',
  'mathEqual',
  'mathNotEqual',
  'cornerTabs',
  'squareTabs',
  'plaqueTabs',
  'chartX',
  'chartStar',
  'chartPlus',
];

const PRESET_SHAPE_SET: ReadonlySet<string> = new Set(PRESET_SHAPE_NAMES);

/** True iff `name` is one of the 187 ECMA-376 preset shapes. */
export const isPresetShapeName = (name: string): boolean => PRESET_SHAPE_SET.has(name);

/** `<a:gd name="..." fmla="..."/>`. Shape guide / formula. */
export interface ShapeGuide {
  name: string;
  fmla: string;
}

/**
 * Adjust point. ECMA-376 calls these out as a Coordinate (a string that
 * may be a literal EMU number or a guide-name reference like `wd2`). We
 * keep them as strings so guide-references survive round-trip.
 */
export interface AdjPoint2D {
  x: string;
  y: string;
}

/** `<a:cxn ang="..."><a:pos x y/></a:cxn>`. */
export interface ConnectionSite {
  /** Angle: either a number (degrees × 60_000) or a guide reference. */
  ang: string;
  pos: AdjPoint2D;
}

/** `<a:ahXY>`-style adjust handle. The polar variant uses the same shape with `ang` in `pos`. */
export interface AdjustHandle {
  /** `xy` (cartesian) or `polar`. */
  kind: 'xy' | 'polar';
  pos: AdjPoint2D;
  /** Range constraints (xy) or radius/angle constraints (polar). */
  gdRefX?: string;
  minX?: string;
  maxX?: string;
  gdRefY?: string;
  minY?: string;
  maxY?: string;
  gdRefR?: string;
  minR?: string;
  maxR?: string;
  gdRefAng?: string;
  minAng?: string;
  maxAng?: string;
}

/** `<a:rect l t r b/>`. Optional bounding rect for custom geometry. */
export interface GuideRect {
  l: string;
  t: string;
  r: string;
  b: string;
}

/** Path-command discriminated union. ECMA-376 §20.1.9.* */
export type PathCommand =
  | { kind: 'moveTo'; pt: Point2D }
  | { kind: 'lnTo'; pt: Point2D }
  | { kind: 'arcTo'; wR: string; hR: string; stAng: string; swAng: string }
  | { kind: 'quadBezTo'; pts: [Point2D, Point2D] }
  | { kind: 'cubicBezTo'; pts: [Point2D, Point2D, Point2D] }
  | { kind: 'close' };

export type PathFill = 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';

/** A single `<a:path>` element inside `<a:pathLst>`. */
export interface GeometryPath {
  w?: number;
  h?: number;
  fill?: PathFill;
  stroke?: boolean;
  extrusionOk?: boolean;
  commands: PathCommand[];
}

/** `<a:prstGeom prst="..."><a:avLst/></a:prstGeom>`. */
export interface PresetGeometry {
  kind: 'preset';
  prst: string;
  avLst?: ShapeGuide[];
}

/** `<a:custGeom>...</a:custGeom>`. */
export interface CustomGeometry {
  kind: 'custom';
  avLst?: ShapeGuide[];
  gdLst?: ShapeGuide[];
  ahLst?: AdjustHandle[];
  cxnLst?: ConnectionSite[];
  rect?: GuideRect;
  pathLst: GeometryPath[];
}

export type Geometry = PresetGeometry | CustomGeometry;

export const makePresetGeometry = (prst: string, avLst?: ShapeGuide[]): PresetGeometry => ({
  kind: 'preset',
  prst,
  ...(avLst && avLst.length > 0 ? { avLst } : {}),
});

export const makeCustomGeometry = (opts: {
  pathLst: GeometryPath[];
  avLst?: ShapeGuide[];
  gdLst?: ShapeGuide[];
  ahLst?: AdjustHandle[];
  cxnLst?: ConnectionSite[];
  rect?: GuideRect;
}): CustomGeometry => ({
  kind: 'custom',
  pathLst: opts.pathLst,
  ...(opts.avLst && opts.avLst.length > 0 ? { avLst: opts.avLst } : {}),
  ...(opts.gdLst && opts.gdLst.length > 0 ? { gdLst: opts.gdLst } : {}),
  ...(opts.ahLst && opts.ahLst.length > 0 ? { ahLst: opts.ahLst } : {}),
  ...(opts.cxnLst && opts.cxnLst.length > 0 ? { cxnLst: opts.cxnLst } : {}),
  ...(opts.rect ? { rect: opts.rect } : {}),
});
