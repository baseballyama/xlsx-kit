// Worksheet-level <scenarios> — Excel's Data → What-If Analysis →
// Scenario Manager. Per ECMA-376 §18.3.1.74 / §18.3.1.41 (inputCells)
// and openpyxl/openpyxl/worksheet/scenario.py.

import type { MultiCellRange } from './cell-range';

/** One <inputCells> entry — a single (cell, override-value) pair. */
export interface ScenarioInputCell {
  /** Single-cell ref, e.g. "B5". */
  ref: string;
  /** Stored as a string on the wire (Excel uses int / float / text format codes via numFmtId). */
  val: string;
  /** Mark this entry as deleted in the scenario history. */
  deleted?: boolean;
  /** Marks an undone change in the scenario history. */
  undone?: boolean;
  /** Number-format index used to display the override. */
  numFmtId?: number;
}

export interface Scenario {
  name: string;
  inputCells: ScenarioInputCell[];
  /** When true, Excel disables editing the scenario unless the workbook protection password is supplied. */
  locked?: boolean;
  /** Hide the scenario from the picker dialog. */
  hidden?: boolean;
  user?: string;
  comment?: string;
}

export interface ScenarioList {
  scenarios: Scenario[];
  /** Index of the currently active scenario. */
  current?: number;
  /** Index of the scenario shown by default. */
  show?: number;
  /** Range that the scenarios change (output cells). */
  sqref?: MultiCellRange;
}

export const makeScenarioInputCell = (opts: ScenarioInputCell): ScenarioInputCell => ({
  ref: opts.ref,
  val: opts.val,
  ...(opts.deleted !== undefined ? { deleted: opts.deleted } : {}),
  ...(opts.undone !== undefined ? { undone: opts.undone } : {}),
  ...(opts.numFmtId !== undefined ? { numFmtId: opts.numFmtId } : {}),
});

export const makeScenario = (opts: Scenario): Scenario => ({
  name: opts.name,
  inputCells: opts.inputCells.slice(),
  ...(opts.locked !== undefined ? { locked: opts.locked } : {}),
  ...(opts.hidden !== undefined ? { hidden: opts.hidden } : {}),
  ...(opts.user !== undefined ? { user: opts.user } : {}),
  ...(opts.comment !== undefined ? { comment: opts.comment } : {}),
});

export const makeScenarioList = (opts: Partial<ScenarioList> = {}): ScenarioList => ({
  scenarios: opts.scenarios?.slice() ?? [],
  ...(opts.current !== undefined ? { current: opts.current } : {}),
  ...(opts.show !== undefined ? { show: opts.show } : {}),
  ...(opts.sqref !== undefined ? { sqref: opts.sqref } : {}),
});
