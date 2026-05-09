// Sheet-protection model. (without password hashing — saltValue / spinCount /
// algorithmName / hashValue round-trip verbatim, but no helper to compute them
// yet).
//
// Excel uses the listed booleans inversely: `true` typically means "users CAN
// do this even when the sheet is locked" (e.g. `formatCells: true` lets people
// change cell formatting on a protected sheet). The only universally meaningful
// field is `sheet: true`, which actually enables the lock.

export interface SheetProtection {
  /** Master toggle — when true the sheet is protected. */
  sheet?: boolean;
  /** Allow operations on drawing objects when sheet is protected. */
  objects?: boolean;
  /** Allow operations on scenarios when sheet is protected. */
  scenarios?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  insertHyperlinks?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  pivotTables?: boolean;

  // Password-protection fields. Round-trip only — computing a fresh hash from a
  // plaintext password lives behind a future helper.
  /** Base-64 salt for the password hash. */
  saltValue?: string;
  /** Number of hash iterations. */
  spinCount?: number;
  /** Hash algorithm name, e.g. "SHA-512". */
  algorithmName?: string;
  /** Base-64 hashed password. */
  hashValue?: string;
}

export const makeSheetProtection = (opts: SheetProtection = {}): SheetProtection => {
  const out: SheetProtection = {};
  for (const k of [
    'sheet',
    'objects',
    'scenarios',
    'formatCells',
    'formatColumns',
    'formatRows',
    'insertColumns',
    'insertRows',
    'insertHyperlinks',
    'deleteColumns',
    'deleteRows',
    'selectLockedCells',
    'selectUnlockedCells',
    'sort',
    'autoFilter',
    'pivotTables',
  ] as const) {
    if (opts[k] !== undefined) out[k] = opts[k];
  }
  if (opts.saltValue !== undefined) out.saltValue = opts.saltValue;
  if (opts.spinCount !== undefined) out.spinCount = opts.spinCount;
  if (opts.algorithmName !== undefined) out.algorithmName = opts.algorithmName;
  if (opts.hashValue !== undefined) out.hashValue = opts.hashValue;
  return out;
};

// ---- Worksheet ergonomic helpers ----------------------------------------

import type { Worksheet } from './worksheet';

/**
 * Excel's "Protect Sheet" defaults — when you click the dialog without changing
 * any checkbox, it locks structure but allows the listed actions. This matches
 * Excel's wire form (sheet=1 + the listed flags left at their defaults).
 */
const PROTECT_SHEET_DEFAULTS: SheetProtection = Object.freeze({
  sheet: true,
  objects: true,
  scenarios: true,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  selectLockedCells: false,
  selectUnlockedCells: false,
  sort: false,
  autoFilter: false,
  pivotTables: false,
});

/**
 * Lock a worksheet with Excel's "Protect Sheet" defaults. Pass `overrides` to
 * allow specific actions while otherwise locked (e.g. `{ sort: true,
 * autoFilter: true }` for "allow sort + filter on locked sheet"). Password-hash
 * fields can be supplied as a quad (algorithmName / hashValue / saltValue /
 * spinCount); plaintext passwords are out of scope until the D-tier hashing
 * helper lands.
 */
export const protectSheet = (
  ws: Worksheet,
  overrides: Partial<SheetProtection> = {},
): SheetProtection => {
  ws.sheetProtection = { ...PROTECT_SHEET_DEFAULTS, ...overrides };
  return ws.sheetProtection;
};

/** Drop the typed sheet-protection record. */
export const unprotectSheet = (ws: Worksheet): void => {
  delete (ws as { sheetProtection?: SheetProtection }).sheetProtection;
};

/** Quick-lock helper that mirrors Excel's "Allow users to edit ranges → Protect Sheet" defaults. */
export const isSheetProtected = (ws: Worksheet): boolean => ws.sheetProtection?.sheet === true;
