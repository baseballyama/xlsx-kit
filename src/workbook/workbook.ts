// Workbook root model. Per docs/plan/04-core-model.md §4.1 / §4.2 /
// docs/plan/01-architecture.md §5 the Workbook is a plain mutable
// object the user composes via free functions. The Stylesheet pool
// is held inline so styling operations don't need a side channel.

import type { Chartsheet } from '../chartsheet/chartsheet';
import { makeChartsheet } from '../chartsheet/chartsheet';
import { makeAbsoluteAnchor } from '../drawing/anchor';
import { type ChartReference, makeChartDrawingItem, makeDrawing } from '../drawing/drawing';
import type { CoreProperties } from '../packaging/core';
import type { CustomProperties } from '../packaging/custom';
import type { ExtendedProperties } from '../packaging/extended';
import type { Stylesheet } from '../styles/stylesheet';
import { makeStylesheet } from '../styles/stylesheet';
import { OpenXmlSchemaError } from '../utils/exceptions';
import type { Worksheet } from '../worksheet/worksheet';
import { makeWorksheet } from '../worksheet/worksheet';

export type SheetState = 'visible' | 'hidden' | 'veryHidden';

/**
 * Discriminated union over the two kinds of sheet a workbook can host.
 * Both variants share `title` (via `sheet.title`) plus the OOXML
 * `sheetId` and `state` attributes; consumers narrow on `kind` to reach
 * the worksheet- vs chartsheet-specific data.
 */
export type SheetRef =
  | { kind: 'worksheet'; sheet: Worksheet; sheetId: number; state: SheetState; rId?: string }
  | { kind: 'chartsheet'; sheet: Chartsheet; sheetId: number; state: SheetState; rId?: string };

export interface Workbook {
  sheets: SheetRef[];
  /** Index into `sheets` of the sheet shown when Excel opens the file. */
  activeSheetIndex: number;
  /** Style pool; cells reference its cellXfs by index. */
  styles: Stylesheet;
  /** Date1904 mode toggles between Excel's two epoch systems. */
  date1904: boolean;
  /** Document properties (docProps/core.xml), typically auto-filled on save. */
  properties?: CoreProperties;
  appProperties?: ExtendedProperties;
  customProperties?: CustomProperties;
  /** Author display names, shared between threaded comments. */
  authors: string[];
  /** Workbook + sheet-scope defined names (named ranges, print areas etc). */
  definedNames: import('./defined-names').DefinedName[];
  /**
   * Raw `xl/theme/theme1.xml` payload kept verbatim across read → write.
   * The theme XML is large and seldom edited by writers; we just shuttle it.
   */
  themeXml?: Uint8Array;
  /**
   * `xl/vbaProject.bin` payload (macro-enabled workbooks). Round-tripped
   * byte-identical when present; the writer also promotes the workbook
   * Override to `vnd.ms-excel.sheet.macroEnabled.main+xml`.
   */
  vbaProject?: Uint8Array;
  /** `xl/vbaProjectSignature.bin` payload, when the macros are signed. */
  vbaSignature?: Uint8Array;
  /**
   * Pass-through bytes for parts we don't model (pivot tables, ActiveX
   * controls, OLE embeddings, customUI ribbons, customXml items …).
   * Keys are archive-relative paths; values are the raw bytes the loader
   * pulled out of the zip and the writer pushes back in unchanged.
   */
  passthrough?: Map<string, Uint8Array>;
  /**
   * Override content type per pass-through path. Excel uses these in
   * `[Content_Types].xml` so manifest validation stays intact across
   * round-trips. Paths without an explicit override fall back to the
   * archive Default extension.
   */
  passthroughContentTypes?: Map<string, string>;
  /**
   * Top-level `<workbook>` children that aren't `<sheets>` or
   * `<definedNames>` (e.g. `<fileVersion>`, `<workbookPr>`,
   * `<bookViews>`, `<calcPr>`, `<pivotCaches>`, `<extLst>`). Captured
   * verbatim so re-saving keeps Excel-rendering fidelity for things we
   * don't model. Split into the two halves the writer needs: anything
   * before `<sheets>` is emitted ahead of the `<sheets>` element, the
   * rest after `<definedNames>`.
   */
  workbookXmlExtras?: {
    beforeSheets: import('../xml/tree').XmlNode[];
    afterSheets: import('../xml/tree').XmlNode[];
  };
  /**
   * `<workbookProtection>` — locks structure / window / revision
   * tracking with the modern hash quad or the legacy 16-bit hash.
   * Round-tripped verbatim; password hashing helpers come later.
   */
  workbookProtection?: import('./protection').WorkbookProtection;
  /**
   * `<bookViews>` — the workbook's window/tab-strip presets. Most
   * workbooks have a single entry whose `firstSheet` / `activeTab`
   * drive the tab the user sees first. Stored as an array because
   * Excel allows multiple views (rare).
   */
  bookViews?: import('./views').WorkbookView[];
  /**
   * Workbook-level rels that don't match a modeled type. Re-emitted with
   * their original Id so captured `<pivotCaches r:id="…"/>` etc. still
   * resolve after a round-trip.
   */
  workbookRelsExtras?: ReadonlyArray<{ id: string; type: string; target: string }>;
  /**
   * Original rIds for the modeled non-sheet workbook rels so a captured
   * extras XML referencing one of them still resolves after re-save.
   */
  workbookRelOriginalIds?: {
    sharedStrings?: string;
    styles?: string;
    theme?: string;
    vbaProject?: string;
  };
}

/** Build an empty Workbook ready to host worksheets. */
export function createWorkbook(opts?: { date1904?: boolean }): Workbook {
  return {
    sheets: [],
    activeSheetIndex: 0,
    styles: makeStylesheet(),
    date1904: opts?.date1904 ?? false,
    authors: [],
    definedNames: [],
  };
}

const validateUniqueTitle = (wb: Workbook, title: string): void => {
  if (typeof title !== 'string' || title.length === 0 || title.length > 31) {
    throw new OpenXmlSchemaError(`Worksheet title must be 1..31 chars; got "${title}"`);
  }
  for (const s of wb.sheets) {
    if (s.sheet.title === title) {
      throw new OpenXmlSchemaError(`Worksheet title "${title}" is already in use`);
    }
  }
};

const allocateSheetId = (wb: Workbook): number => {
  // sheetId is 1-based and unique. Allocate the smallest unused integer.
  const used = new Set<number>();
  for (const s of wb.sheets) used.add(s.sheetId);
  let n = 1;
  while (used.has(n)) n++;
  return n;
};

/** Add a Worksheet to the Workbook. Returns the sheet for further population. */
export function addWorksheet(wb: Workbook, title: string, opts?: { index?: number; state?: SheetState }): Worksheet {
  validateUniqueTitle(wb, title);
  const sheet = makeWorksheet(title);
  const ref: SheetRef = {
    kind: 'worksheet',
    sheet,
    sheetId: allocateSheetId(wb),
    state: opts?.state ?? 'visible',
  };
  if (opts?.index === undefined) {
    wb.sheets.push(ref);
  } else {
    if (opts.index < 0 || opts.index > wb.sheets.length) {
      throw new OpenXmlSchemaError(`addWorksheet: index ${opts.index} out of range`);
    }
    wb.sheets.splice(opts.index, 0, ref);
  }
  return sheet;
}

/** Look up a Worksheet by title. Returns undefined for missing names or chartsheets. */
export function getSheet(wb: Workbook, title: string): Worksheet | undefined {
  for (const s of wb.sheets) {
    if (s.kind === 'worksheet' && s.sheet.title === title) return s.sheet;
  }
  return undefined;
}

/** Look up a Worksheet by index in the sheets array. Returns undefined for chartsheet slots. */
export function getSheetByIndex(wb: Workbook, idx: number): Worksheet | undefined {
  const ref = wb.sheets[idx];
  return ref?.kind === 'worksheet' ? ref.sheet : undefined;
}

/** Look up a Chartsheet by title. Returns undefined for missing names or worksheets. */
export function getChartsheet(wb: Workbook, title: string): Chartsheet | undefined {
  for (const s of wb.sheets) {
    if (s.kind === 'chartsheet' && s.sheet.title === title) return s.sheet;
  }
  return undefined;
}

/** Add a Chartsheet to the Workbook. Returns the chartsheet for further population. */
export function addChartsheet(
  wb: Workbook,
  title: string,
  opts?: { index?: number; state?: SheetState; chart?: ChartReference },
): Chartsheet {
  validateUniqueTitle(wb, title);
  const cs = makeChartsheet(title);
  // When the caller supplies a ChartReference, wrap it in a single-anchor
  // drawing so the writer emits xl/drawings/drawingN.xml + chart part.
  if (opts?.chart) {
    // The drawing wraps the chart in a single absoluteAnchor sized to a
    // standard A4-landscape page (Excel re-flows on open if needed).
    cs.drawing = makeDrawing([
      makeChartDrawingItem(makeAbsoluteAnchor({ x: 0, y: 0, cx: 9144000, cy: 6858000 }), opts.chart),
    ]);
  }
  const ref: SheetRef = {
    kind: 'chartsheet',
    sheet: cs,
    sheetId: allocateSheetId(wb),
    state: opts?.state ?? 'visible',
  };
  if (opts?.index === undefined) {
    wb.sheets.push(ref);
  } else {
    if (opts.index < 0 || opts.index > wb.sheets.length) {
      throw new OpenXmlSchemaError(`addChartsheet: index ${opts.index} out of range`);
    }
    wb.sheets.splice(opts.index, 0, ref);
  }
  return cs;
}

/** All worksheet titles, in display order. */
export function sheetNames(wb: Workbook): string[] {
  return wb.sheets.map((s) => s.sheet.title);
}

/** Remove a sheet by title. No-op if the title is not registered. */
export function removeSheet(wb: Workbook, title: string): void {
  const i = wb.sheets.findIndex((s) => s.sheet.title === title);
  if (i < 0) return;
  wb.sheets.splice(i, 1);
  // Clamp activeSheetIndex within bounds.
  if (wb.activeSheetIndex >= wb.sheets.length) {
    wb.activeSheetIndex = Math.max(0, wb.sheets.length - 1);
  }
}

/** Set the active sheet by title; throws on unknown title. */
export function setActiveSheet(wb: Workbook, title: string): void {
  const i = wb.sheets.findIndex((s) => s.sheet.title === title);
  if (i < 0) throw new OpenXmlSchemaError(`setActiveSheet: no sheet named "${title}"`);
  wb.activeSheetIndex = i;
}

/** Currently active sheet (worksheet only), or undefined if the active slot is empty or a chartsheet. */
export function getActiveSheet(wb: Workbook): Worksheet | undefined {
  const ref = wb.sheets[wb.activeSheetIndex];
  return ref?.kind === 'worksheet' ? ref.sheet : undefined;
}

/** Read-only view onto the customXml/* pass-through parts. */
export function listCustomXmlParts(wb: Workbook): Array<{ path: string; content: Uint8Array }> {
  if (!wb.passthrough) return [];
  const out: Array<{ path: string; content: Uint8Array }> = [];
  for (const [path, content] of wb.passthrough) {
    if (path.startsWith('customXml/')) out.push({ path, content });
  }
  return out;
}

/**
 * JSON.stringify replacer that drops the Stylesheet's internal dedup
 * Maps. Use as `JSON.stringify(workbook, jsonReplacer)` when the
 * workbook needs to round-trip through plain JSON (tests, debug
 * dumps). The dedup maps are reconstructed lazily on first add.
 */
export function jsonReplacer(_key: string, value: unknown): unknown {
  if (value instanceof Map) {
    return { __map__: [...value.entries()] };
  }
  return value;
}

/** Companion reviver for {@link jsonReplacer}. */
export function jsonReviver(_key: string, value: unknown): unknown {
  if (typeof value === 'object' && value !== null && Array.isArray((value as { __map__?: unknown[] }).__map__)) {
    return new Map((value as { __map__: Array<[unknown, unknown]> }).__map__);
  }
  return value;
}
