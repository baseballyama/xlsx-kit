// Workbook root model. Per docs/plan/04-core-model.md §4.1 / §4.2 /
// docs/plan/01-architecture.md §5 the Workbook is a plain mutable
// object the user composes via free functions. The Stylesheet pool
// is held inline so styling operations don't need a side channel.

import type { CoreProperties } from '../packaging/core';
import type { CustomProperties } from '../packaging/custom';
import type { ExtendedProperties } from '../packaging/extended';
import type { Stylesheet } from '../styles/stylesheet';
import { makeStylesheet } from '../styles/stylesheet';
import { OpenXmlSchemaError } from '../utils/exceptions';
import type { Worksheet } from '../worksheet/worksheet';
import { makeWorksheet } from '../worksheet/worksheet';

export type SheetState = 'visible' | 'hidden' | 'veryHidden';

export interface SheetRef {
  /** Kind tag — only `worksheet` for now; chartsheet lands in phase 6. */
  kind: 'worksheet';
  sheet: Worksheet;
  /** OOXML sheetId attribute (1-based, unique within the workbook). */
  sheetId: number;
  state: SheetState;
}

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

/** Look up a Worksheet by title. Returns undefined if absent. */
export function getSheet(wb: Workbook, title: string): Worksheet | undefined {
  for (const s of wb.sheets) if (s.sheet.title === title) return s.sheet;
  return undefined;
}

/** Look up a Worksheet by index in the sheets array. */
export function getSheetByIndex(wb: Workbook, idx: number): Worksheet | undefined {
  return wb.sheets[idx]?.sheet;
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

/** Currently active worksheet, or undefined if the workbook is empty. */
export function getActiveSheet(wb: Workbook): Worksheet | undefined {
  return wb.sheets[wb.activeSheetIndex]?.sheet;
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
