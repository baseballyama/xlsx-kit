// Workbook-level <fileSharing>. Per ECMA-376 §18.2.12.
//
// Carries the workbook's read-only / write-protection settings (Save
// As → Tools → General Options → "Modify password" / "Read-only
// recommended"). The hash quad mirrors sheetProtection / workbookProtection.

export interface FileSharing {
  /** Mark the workbook as "Read-only recommended" — Excel pops a dialog on open. */
  readOnlyRecommended?: boolean;
  /** Author name attached to the read/write password. */
  userName?: string;
  /** Legacy 16-bit hex hash of the reservation password. */
  reservationPassword?: string;
  /** Modern hash quad — algorithmName + hashValue + saltValue + spinCount. */
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
}

export const makeFileSharing = (opts: FileSharing = {}): FileSharing => ({ ...opts });
