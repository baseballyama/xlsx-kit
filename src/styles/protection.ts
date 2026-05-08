// Cell protection value object. Mirrors openpyxl/openpyxl/styles/protection.py.
//
// Protection only matters when the worksheet itself is protected; until
// then it's metadata that round-trips with the cell's style.

export interface Protection {
  /** When true (the default), the cell can't be edited if the sheet is locked. */
  readonly locked?: boolean;
  /** When true, the formula bar hides this cell's contents on a protected sheet. */
  readonly hidden?: boolean;
}

export function makeProtection(opts: Partial<Protection> = {}): Protection {
  const out: { -readonly [K in keyof Protection]: Protection[K] } = {};
  if (opts.locked !== undefined) out.locked = opts.locked;
  if (opts.hidden !== undefined) out.hidden = opts.hidden;
  return Object.freeze(out);
}

/** Excel's Stylesheet always contains an entry equal to {locked:true, hidden:false}. */
export const DEFAULT_PROTECTION: Protection = makeProtection({ locked: true, hidden: false });
