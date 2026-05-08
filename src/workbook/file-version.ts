// Workbook-level <fileVersion>. Per ECMA-376 §18.2.13.
//
// Carries Microsoft Office app/version metadata that Excel writes
// when it saves; round-tripping these values keeps the file looking
// like Excel's own output to downstream tools that sniff them.

export interface FileVersion {
  /** Application that last saved the workbook ("xl" for Excel). */
  appName?: string;
  /** Build number of the last editor (e.g. "7.5210"). */
  lastEdited?: string;
  /** Build number of the lowest editor (oldest Excel that touched the file). */
  lowestEdited?: string;
  /** Internal "rolled-up build" number. */
  rupBuild?: string;
  /** GUID identifying the file content (Excel uses it to detect re-saves). */
  codeName?: string;
}

export const makeFileVersion = (opts: FileVersion = {}): FileVersion => ({ ...opts });
