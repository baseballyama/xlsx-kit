// Workbook-level <fileRecoveryPr>. Per ECMA-376 §18.2.11.
//
// Excel writes this element after an autorecover sequence to mark the
// workbook with the recovery state so subsequent opens can prompt the
// user. Almost always absent in fresh files.

export interface FileRecoveryProperties {
  /** True after an autorecover save — Excel uses it to display the recovery banner. */
  autoRecover?: boolean;
  /** Persisted crash-recovery flag. */
  crashSave?: boolean;
  /** Mark the file as "data extracted from a damaged workbook". */
  dataExtractLoad?: boolean;
  /** Workbook was repaired during load. */
  repairLoad?: boolean;
}

export const makeFileRecoveryProperties = (
  opts: FileRecoveryProperties = {},
): FileRecoveryProperties => ({ ...opts });
