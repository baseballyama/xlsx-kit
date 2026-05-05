// Sheet-protection model. Per docs/plan/13-full-excel-coverage.md §B5
// (without password hashing — saltValue / spinCount / algorithmName /
// hashValue round-trip verbatim, but no helper to compute them yet).
//
// Excel uses the listed booleans inversely: `true` typically means
// "users CAN do this even when the sheet is locked" (e.g. `formatCells:
// true` lets people change cell formatting on a protected sheet). The
// only universally meaningful field is `sheet: true`, which actually
// enables the lock.

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

  // Password-protection fields. Round-trip only — computing a fresh
  // hash from a plaintext password lives behind a future helper (see
  // docs/plan/13 §D for the hashing track).
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
