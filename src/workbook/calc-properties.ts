// Workbook-level <calcPr>. Per ECMA-376 §18.2.2.
//
// `calcId` is a build identifier Excel uses to decide whether to force
// a recalc when re-opening; modern Excel treats any value as "use
// what's there". The other attrs cover Excel's calculation options
// (Tools → Options → Formulas).

export type CalcMode = 'manual' | 'auto' | 'autoNoTable';
export type RefMode = 'A1' | 'R1C1';

export interface CalcProperties {
  /** Excel build/calc-engine identifier; OpenOffice sets 191621, Excel 2016+ 162913 etc. */
  calcId?: number;
  calcMode?: CalcMode;
  /** Force a full recalc when the workbook is loaded. Excel default `true` for safety. */
  fullCalcOnLoad?: boolean;
  refMode?: RefMode;
  /** Allow circular references via iterative calculation. */
  iterate?: boolean;
  iterateCount?: number;
  iterateDelta?: number;
  /** Use full 15-digit precision (vs. displayed precision). */
  fullPrecision?: boolean;
  calcCompleted?: boolean;
  /** Run a recalc on save. */
  calcOnSave?: boolean;
  /** Use multi-threaded calculation. */
  concurrentCalc?: boolean;
  concurrentManualCount?: number;
  /** Force a full recalc on next interaction. */
  forceFullCalc?: boolean;
}

export const makeCalcProperties = (opts: CalcProperties = {}): CalcProperties => ({ ...opts });
