// Workbook-level defined names. Per docs/plan/07-rich-features.md §3
// + the OOXML schema's `<definedName>` element.
//
// A defined name binds an identifier to a formula-style value
// (`'Sheet 1'!$A$1:$B$10`, `SUM(A:A)`, etc). Workbook-scope names omit
// `localSheetId`; sheet-scope names use the 0-based sheet index. Excel
// reserves a handful of names with the `_xlnm.` prefix for built-in
// uses (Print_Area, Print_Titles, Sheet_Title, etc) — those round-trip
// here as plain DefinedName entries since the value semantics are the
// same.

export interface DefinedName {
  /** Identifier — `_xlnm.Print_Area` for built-ins, otherwise user-chosen. */
  name: string;
  /** The formula expression the name points at. */
  value: string;
  /** 0-based sheet index for sheet-scope names; undefined → workbook-scope. */
  scope?: number;
  /** Hidden from the Name Manager when true. */
  hidden?: boolean;
  /** Optional human-readable description. */
  comment?: string;
}

export function makeDefinedName(opts: Partial<DefinedName> & { name: string; value: string }): DefinedName {
  return {
    name: opts.name,
    value: opts.value,
    ...(opts.scope !== undefined ? { scope: opts.scope } : {}),
    ...(opts.hidden !== undefined ? { hidden: opts.hidden } : {}),
    ...(opts.comment !== undefined ? { comment: opts.comment } : {}),
  };
}

// ---- Workbook ergonomic helpers -----------------------------------------

import type { Workbook } from './workbook';

/**
 * Add a workbook-scope or sheet-scope defined name. If a defined name
 * with the same `name` (and `scope`) already exists, it's replaced —
 * Excel allows one workbook-scope and one per-sheet-scope name, but
 * not two with the same scope. Returns the resulting `DefinedName`.
 */
export const addDefinedName = (
  wb: Workbook,
  opts: Partial<DefinedName> & { name: string; value: string },
): DefinedName => {
  const dn = makeDefinedName(opts);
  // Replace any existing entry with the same name + scope.
  const idx = wb.definedNames.findIndex((d) => d.name === dn.name && d.scope === dn.scope);
  if (idx >= 0) {
    wb.definedNames[idx] = dn;
  } else {
    wb.definedNames.push(dn);
  }
  return dn;
};

/** Look up a defined name by identifier and (optional) sheet scope. */
export const getDefinedName = (
  wb: Workbook,
  name: string,
  scope?: number,
): DefinedName | undefined => wb.definedNames.find((d) => d.name === name && d.scope === scope);

/**
 * Remove a defined name by identifier + scope. Returns true if any
 * entry was removed.
 */
export const removeDefinedName = (wb: Workbook, name: string, scope?: number): boolean => {
  const idx = wb.definedNames.findIndex((d) => d.name === name && d.scope === scope);
  if (idx < 0) return false;
  wb.definedNames.splice(idx, 1);
  return true;
};

/**
 * Read-only snapshot of every defined name. Pass `{ scope }` to
 * narrow to workbook-scope (`scope: undefined`) or one specific
 * sheet (`scope: 0`) — omit the option entirely to list all.
 */
export const listDefinedNames = (
  wb: Workbook,
  opts: { scope?: number | 'workbook' | 'all' } = {},
): ReadonlyArray<DefinedName> => {
  const scope = opts.scope ?? 'all';
  if (scope === 'all') return wb.definedNames;
  if (scope === 'workbook') return wb.definedNames.filter((d) => d.scope === undefined);
  return wb.definedNames.filter((d) => d.scope === scope);
};

/**
 * Bulk-remove every defined name matching `predicate`. Returns the
 * count removed. Mirrors {@link removeDataValidations} on worksheets.
 */
export const removeDefinedNames = (
  wb: Workbook,
  predicate: (d: DefinedName) => boolean,
): number => {
  const before = wb.definedNames.length;
  wb.definedNames = wb.definedNames.filter((d) => !predicate(d));
  return before - wb.definedNames.length;
};

/**
 * Define the print-area for a given sheet. Excel uses the built-in
 * `_xlnm.Print_Area` defined name with sheet scope.
 */
export const setPrintArea = (wb: Workbook, sheetIndex: number, ref: string): DefinedName => {
  return addDefinedName(wb, {
    name: '_xlnm.Print_Area',
    value: ref,
    scope: sheetIndex,
  });
};

/**
 * Define print-title rows / columns on a sheet. Excel uses the
 * `_xlnm.Print_Titles` defined name. Pass `rows` ("$1:$1") to repeat
 * row 1 on every printed page; `cols` ("$A:$A") to repeat column A.
 */
export const setPrintTitles = (
  wb: Workbook,
  sheetIndex: number,
  opts: { rows?: string; cols?: string; sheetName: string },
): DefinedName => {
  const parts: string[] = [];
  // The wire form is "Sheet!$1:$1,Sheet!$A:$A"; both refs share the
  // sheet prefix.
  if (opts.cols !== undefined) parts.push(`'${opts.sheetName}'!${opts.cols}`);
  if (opts.rows !== undefined) parts.push(`'${opts.sheetName}'!${opts.rows}`);
  if (parts.length === 0) {
    throw new Error('setPrintTitles: at least one of rows or cols must be set');
  }
  return addDefinedName(wb, {
    name: '_xlnm.Print_Titles',
    value: parts.join(','),
    scope: sheetIndex,
  });
};
