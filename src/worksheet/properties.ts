// Worksheet `<sheetPr>` properties. Per docs/plan/13-full-excel-coverage.md
// §B7 (sheet view 拡張). Mirrors openpyxl/openpyxl/worksheet/properties.py.
//
// Promoting sheetPr out of `bodyExtras.beforeSheetData` gives consumers a
// typed handle for the most common fields (`tabColor`, `codeName`) and
// for the two child elements that drive Excel's outline / page-setup
// behaviour. The full set is modeled so a load → save round-trip
// preserves the element verbatim — niche sync* / transition* attrs
// included.

import type { Color } from '../styles/colors';

export interface OutlineProperties {
  applyStyles?: boolean;
  /** "Summary rows below detail" — Excel default true. */
  summaryBelow?: boolean;
  /** "Summary columns to right of detail" — Excel default true. */
  summaryRight?: boolean;
  showOutlineSymbols?: boolean;
}

export interface PageSetupProperties {
  autoPageBreaks?: boolean;
  fitToPage?: boolean;
}

export interface SheetProperties {
  /** VBA codeName for the sheet ("Sheet1" by default but localizable). */
  codeName?: string;
  enableFormatConditionsCalculation?: boolean;
  /** Whether the sheet has an active filter set up. */
  filterMode?: boolean;
  /** Whether the sheet is published to a SharePoint Excel Services list. */
  published?: boolean;
  /** Sync* attrs control multi-sheet scroll synchronisation. */
  syncHorizontal?: boolean;
  syncRef?: string;
  syncVertical?: boolean;
  /** "Lotus 1-2-3 transition" toggles — almost never seen in modern files. */
  transitionEvaluation?: boolean;
  transitionEntry?: boolean;
  /** Tab strip colour for this sheet (the coloured stripe at the bottom). */
  tabColor?: Color;
  outlinePr?: OutlineProperties;
  pageSetUpPr?: PageSetupProperties;
}

export const makeSheetProperties = (opts: SheetProperties = {}): SheetProperties => {
  const out: SheetProperties = {};
  if (opts.codeName !== undefined) out.codeName = opts.codeName;
  if (opts.enableFormatConditionsCalculation !== undefined)
    out.enableFormatConditionsCalculation = opts.enableFormatConditionsCalculation;
  if (opts.filterMode !== undefined) out.filterMode = opts.filterMode;
  if (opts.published !== undefined) out.published = opts.published;
  if (opts.syncHorizontal !== undefined) out.syncHorizontal = opts.syncHorizontal;
  if (opts.syncRef !== undefined) out.syncRef = opts.syncRef;
  if (opts.syncVertical !== undefined) out.syncVertical = opts.syncVertical;
  if (opts.transitionEvaluation !== undefined) out.transitionEvaluation = opts.transitionEvaluation;
  if (opts.transitionEntry !== undefined) out.transitionEntry = opts.transitionEntry;
  if (opts.tabColor !== undefined) out.tabColor = opts.tabColor;
  if (opts.outlinePr !== undefined) out.outlinePr = opts.outlinePr;
  if (opts.pageSetUpPr !== undefined) out.pageSetUpPr = opts.pageSetUpPr;
  return out;
};
