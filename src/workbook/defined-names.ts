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
