// Worksheet-level <protectedRanges>. Per ECMA-376 §18.3.1.69.
//
// Excel's "Allow Edit Ranges" dialog (Review → Allow Edit Ranges).
// Each entry whitelists a specific range to be editable while the
// sheet is otherwise protected. Round-tripped verbatim — no password
// hashing helper yet.

import type { MultiCellRange } from './cell-range';

export interface ProtectedRange {
  /** Range to expose for editing while the sheet is protected. */
  sqref: MultiCellRange;
  /** Display name shown in the dialog. */
  name: string;
  /** Legacy 16-bit hex password. */
  password?: string;
  /** Optional security descriptor (Windows ACL string). */
  securityDescriptor?: string;
  // Modern hash quad — round-tripped verbatim.
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
}

export const makeProtectedRange = (
  opts: Partial<ProtectedRange> & { sqref: MultiCellRange; name: string },
): ProtectedRange => ({
  sqref: opts.sqref,
  name: opts.name,
  ...(opts.password !== undefined ? { password: opts.password } : {}),
  ...(opts.securityDescriptor !== undefined ? { securityDescriptor: opts.securityDescriptor } : {}),
  ...(opts.algorithmName !== undefined ? { algorithmName: opts.algorithmName } : {}),
  ...(opts.hashValue !== undefined ? { hashValue: opts.hashValue } : {}),
  ...(opts.saltValue !== undefined ? { saltValue: opts.saltValue } : {}),
  ...(opts.spinCount !== undefined ? { spinCount: opts.spinCount } : {}),
});
