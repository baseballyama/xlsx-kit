// Workbook-level <workbookProtection>. Per ECMA-376 §18.2.29 and (workbook
// side).
//
// Two parallel password-hash quadruples cover Excel's "Protect Workbook"
// dialog: one set locks structure / window resize, the other locks the
// revision-tracking history. The wire-form attrs round-trip verbatim —
// computing a fresh hash from a plaintext password lives behind a future helper
// (D-tier in the roadmap).
//
// Note: the non-hash `workbookPassword` / `revisionsPassword` attrs are the
// legacy Excel 97/2000 hex hash form. Modern files use the `*HashValue` +
// `*SaltValue` + `*SpinCount` + `*AlgorithmName` quad instead.

export interface WorkbookProtection {
  /** Legacy 16-bit hex hash of the workbook password ("CC1A" etc.). */
  workbookPassword?: string;
  workbookPasswordCharacterSet?: string;
  workbookAlgorithmName?: string;
  workbookHashValue?: string;
  workbookSaltValue?: string;
  workbookSpinCount?: number;
  /** Legacy 16-bit hex hash of the revisions-tracking password. */
  revisionsPassword?: string;
  revisionsPasswordCharacterSet?: string;
  revisionsAlgorithmName?: string;
  revisionsHashValue?: string;
  revisionsSaltValue?: string;
  revisionsSpinCount?: number;
  /** Lock add/delete/move/rename/hide of sheets. */
  lockStructure?: boolean;
  /** Lock the workbook window size and position. */
  lockWindows?: boolean;
  /** Lock revision tracking — enabled with the "Track Changes" feature. */
  lockRevision?: boolean;
}

export const makeWorkbookProtection = (opts: WorkbookProtection = {}): WorkbookProtection => {
  const out: WorkbookProtection = {};
  for (const k of [
    'workbookPassword',
    'workbookPasswordCharacterSet',
    'workbookAlgorithmName',
    'workbookHashValue',
    'workbookSaltValue',
    'revisionsPassword',
    'revisionsPasswordCharacterSet',
    'revisionsAlgorithmName',
    'revisionsHashValue',
    'revisionsSaltValue',
  ] as const) {
    if (opts[k] !== undefined) out[k] = opts[k];
  }
  if (opts.workbookSpinCount !== undefined) out.workbookSpinCount = opts.workbookSpinCount;
  if (opts.revisionsSpinCount !== undefined) out.revisionsSpinCount = opts.revisionsSpinCount;
  if (opts.lockStructure !== undefined) out.lockStructure = opts.lockStructure;
  if (opts.lockWindows !== undefined) out.lockWindows = opts.lockWindows;
  if (opts.lockRevision !== undefined) out.lockRevision = opts.lockRevision;
  return out;
};

// ---- Workbook ergonomic helpers -----------------------------------------

import type { Workbook } from './workbook';

/**
 * Lock the workbook with Excel's "Protect Workbook → Structure" default
 * (lockStructure=true). Pass `overrides` to also lock windows /
 * revision-tracking, or to attach a password-hash quad. Plaintext password
 * support is deferred until the D-tier hashing helper lands.
 */
export const protectWorkbook = (
  wb: Workbook,
  overrides: Partial<WorkbookProtection> = {},
): WorkbookProtection => {
  wb.workbookProtection = { lockStructure: true, ...overrides };
  return wb.workbookProtection;
};

/** Drop the workbook-protection record entirely. */
export const unprotectWorkbook = (wb: Workbook): void => {
  delete (wb as { workbookProtection?: WorkbookProtection }).workbookProtection;
};

/** True iff `lockStructure === true`. */
export const isWorkbookProtected = (wb: Workbook): boolean =>
  wb.workbookProtection?.lockStructure === true;
