// Worksheet-level <customSheetViews> — saved per-user view presets.
// Per ECMA-376 §18.3.1.26 / §18.3.1.27.
//
// Each <customSheetView> snapshots a sheet's view state: zoom level,
// gridline / formula / heading toggles, plus its own page-setup block
// and break list. The chartsheet sibling (smaller) has been typed
// separately in src/chartsheet/chartsheet.ts.

import type { HeaderFooter, PageBreak, PageMargins, PageSetup, PrintOptions } from './page-setup';
import type { Pane, Selection, SheetViewMode } from './views';

export type CustomSheetViewState = 'visible' | 'hidden' | 'veryHidden';

export interface CustomSheetView {
  guid: string;
  scale?: number;
  /** 0..64 (legacy palette index) */
  colorId?: number;
  showPageBreaks?: boolean;
  showFormulas?: boolean;
  showGridLines?: boolean;
  showRowCol?: boolean;
  outlineSymbols?: boolean;
  zeroValues?: boolean;
  fitToPage?: boolean;
  /** Print only the print-area selection on this saved view. */
  printArea?: boolean;
  /** AutoFilter is active in this saved view. */
  filter?: boolean;
  showAutoFilter?: boolean;
  /** Hidden rows persist for this saved view. */
  hiddenRows?: boolean;
  hiddenColumns?: boolean;
  state?: CustomSheetViewState;
  filterUnique?: boolean;
  view?: SheetViewMode;
  showRuler?: boolean;
  /** Top-left cell ref shown when this view is restored. */
  topLeftCell?: string;
  /** Inner pane split / freeze. */
  pane?: Pane;
  /** Selection state for this view (one entry per pane). */
  selections?: Selection[];
  rowBreaks?: PageBreak[];
  colBreaks?: PageBreak[];
  pageMargins?: PageMargins;
  printOptions?: PrintOptions;
  pageSetup?: PageSetup;
  headerFooter?: HeaderFooter;
}

export const makeCustomSheetView = (
  opts: Partial<CustomSheetView> & { guid: string },
): CustomSheetView => ({ ...opts });
