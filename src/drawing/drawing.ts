// Spreadsheet drawing data model. Per docs/plan/08-charts-drawings.md §3.2.
//
// A `Drawing` is the per-worksheet `xl/drawings/drawingN.xml` part — a
// list of anchor entries, each carrying a content variant (chart,
// picture, shape, connector, group). Stage-1 implements the chart
// variant as a "rels-only" reference (the full ChartML model lands in
// later iterations); picture / shape / connector / group are reserved
// for later.

import type { ChartSpace } from '../chart/chart';
import type { DrawingAnchor } from './anchor';

/** Reference to a chart part — the chart's drawing-rels rId resolves to xl/charts/chartN.xml. */
export interface ChartReference {
  /** Drawing-rels rId pointing at the chart part. Populated on read; the writer assigns its own. */
  rId?: string;
  /**
   * The actual chart payload. Stage-1 supports BarChart end-to-end; other
   * chart kinds populate this field as their parsers / writers land.
   */
  space?: ChartSpace;
}

export interface DrawingItem {
  anchor: DrawingAnchor;
  content: { kind: 'chart'; chart: ChartReference } | { kind: 'unsupported'; rawTag: string };
}

export interface Drawing {
  items: DrawingItem[];
}

export function makeDrawing(items: DrawingItem[] = []): Drawing {
  return { items };
}

export function makeChartDrawingItem(anchor: DrawingAnchor, chart: ChartReference = {}): DrawingItem {
  return { anchor, content: { kind: 'chart', chart } };
}
