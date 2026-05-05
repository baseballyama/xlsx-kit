// Conditional formatting. Per docs/plan/07-rich-features.md §6.
//
// Stage-1 covers the **value-based** rule kinds (cellIs / expression /
// top10 / aboveAverage / containsText family / containsBlanks family /
// duplicateValues / uniqueValues / timePeriod). The visual rule kinds
// — colorScale / dataBar / iconSet — round-trip as opaque inner XML
// (`innerXml` field) so the data survives a save / load cycle without
// our needing to model cfvo / colors / iconSets fully.

import type { MultiCellRange } from './cell-range';

export type ConditionalFormattingRuleType =
  | 'expression'
  | 'cellIs'
  | 'colorScale'
  | 'dataBar'
  | 'iconSet'
  | 'top10'
  | 'aboveAverage'
  | 'uniqueValues'
  | 'duplicateValues'
  | 'containsText'
  | 'notContainsText'
  | 'beginsWith'
  | 'endsWith'
  | 'containsBlanks'
  | 'notContainsBlanks'
  | 'containsErrors'
  | 'notContainsErrors'
  | 'timePeriod';

export type CellIsOperator =
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'greaterThan'
  | 'between'
  | 'notBetween';

export type TextOperator = 'containsText' | 'notContains' | 'beginsWith' | 'endsWith';

export type TimePeriod =
  | 'today'
  | 'yesterday'
  | 'tomorrow'
  | 'last7Days'
  | 'thisMonth'
  | 'lastMonth'
  | 'nextMonth'
  | 'thisWeek'
  | 'lastWeek'
  | 'nextWeek';

export interface ConditionalFormattingRule {
  /** Wire-level rule kind. */
  type: ConditionalFormattingRuleType;
  /** 1-based priority — Excel evaluates lower priority first. */
  priority: number;
  /** Index into Stylesheet.dxfs for the cell-format applied when the rule fires. */
  dxfId?: number;
  /** Stop evaluating subsequent rules on the same cell when this rule matches. */
  stopIfTrue?: boolean;
  /** cellIs operator. */
  operator?: CellIsOperator | TextOperator | string;
  /** Comparison string for the contains-text family. */
  text?: string;
  /** top10: rank percentage flag. */
  percent?: boolean;
  /** top10: bottom-N flag (false = top-N, true = bottom-N). */
  bottom?: boolean;
  /** top10: rank value (defaults to 10). */
  rank?: number;
  /** aboveAverage: aboveAverage="0" → below average. */
  aboveAverage?: boolean;
  /** aboveAverage: equalAverage flag. */
  equalAverage?: boolean;
  /** aboveAverage: stdDev. */
  stdDev?: number;
  /** timePeriod token. */
  timePeriod?: TimePeriod;
  /** 0..3 formula strings — varies by rule type. */
  formulas: string[];
  /**
   * Raw inner XML for colorScale / dataBar / iconSet rules. Stage-1 stores
   * the verbatim child markup so saves round-trip without our needing to
   * model cfvo / colors / iconSets fully.
   */
  innerXml?: string;
}

export interface ConditionalFormatting {
  sqref: MultiCellRange;
  rules: ConditionalFormattingRule[];
  pivot?: boolean;
}

export function makeConditionalFormatting(opts: {
  sqref: MultiCellRange;
  rules?: ConditionalFormattingRule[];
  pivot?: boolean;
}): ConditionalFormatting {
  return {
    sqref: opts.sqref,
    rules: opts.rules ?? [],
    ...(opts.pivot !== undefined ? { pivot: opts.pivot } : {}),
  };
}

export function makeCfRule(
  opts: Partial<ConditionalFormattingRule> & {
    type: ConditionalFormattingRuleType;
    priority: number;
  },
): ConditionalFormattingRule {
  return {
    type: opts.type,
    priority: opts.priority,
    formulas: opts.formulas ?? [],
    ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
    ...(opts.stopIfTrue !== undefined ? { stopIfTrue: opts.stopIfTrue } : {}),
    ...(opts.operator !== undefined ? { operator: opts.operator } : {}),
    ...(opts.text !== undefined ? { text: opts.text } : {}),
    ...(opts.percent !== undefined ? { percent: opts.percent } : {}),
    ...(opts.bottom !== undefined ? { bottom: opts.bottom } : {}),
    ...(opts.rank !== undefined ? { rank: opts.rank } : {}),
    ...(opts.aboveAverage !== undefined ? { aboveAverage: opts.aboveAverage } : {}),
    ...(opts.equalAverage !== undefined ? { equalAverage: opts.equalAverage } : {}),
    ...(opts.stdDev !== undefined ? { stdDev: opts.stdDev } : {}),
    ...(opts.timePeriod !== undefined ? { timePeriod: opts.timePeriod } : {}),
    ...(opts.innerXml !== undefined ? { innerXml: opts.innerXml } : {}),
  };
}

// ---- Worksheet ergonomic builders ---------------------------------------

import type { Worksheet } from './worksheet';
import { parseMultiCellRange } from './cell-range';

const resolveCfSqref = (sqref: MultiCellRange | string): MultiCellRange =>
  typeof sqref === 'string' ? parseMultiCellRange(sqref) : sqref;

const nextCfPriority = (ws: Worksheet): number => {
  let max = 0;
  for (const cf of ws.conditionalFormatting) {
    for (const r of cf.rules) {
      if (r.priority > max) max = r.priority;
    }
  }
  return max + 1;
};

/** Push one CF rule onto the worksheet, wrapping it in a ConditionalFormatting block keyed off `sqref`. */
const pushRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  rule: ConditionalFormattingRule,
): ConditionalFormattingRule => {
  ws.conditionalFormatting.push(makeConditionalFormatting({ sqref: resolveCfSqref(sqref), rules: [rule] }));
  return rule;
};

/**
 * "If cell value [op] formula → apply dxf". Mirrors Excel's "Highlight
 * Cell Rules → ..." UI.
 */
export const addCellIsRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: {
    operator: CellIsOperator;
    formula1: string;
    formula2?: string;
    dxfId?: number;
    stopIfTrue?: boolean;
    priority?: number;
  },
): ConditionalFormattingRule => {
  const formulas = opts.formula2 !== undefined ? [opts.formula1, opts.formula2] : [opts.formula1];
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: 'cellIs',
      priority: opts.priority ?? nextCfPriority(ws),
      operator: opts.operator,
      formulas,
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
      ...(opts.stopIfTrue !== undefined ? { stopIfTrue: opts.stopIfTrue } : {}),
    }),
  );
};

/** Top-N or bottom-N rule. `bottom: true` → bottom-N. `percent: true` → percentile. */
export const addTopNRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: { rank?: number; bottom?: boolean; percent?: boolean; dxfId?: number; priority?: number },
): ConditionalFormattingRule => {
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: 'top10',
      priority: opts.priority ?? nextCfPriority(ws),
      formulas: [],
      ...(opts.rank !== undefined ? { rank: opts.rank } : {}),
      ...(opts.bottom !== undefined ? { bottom: opts.bottom } : {}),
      ...(opts.percent !== undefined ? { percent: opts.percent } : {}),
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
    }),
  );
};

/** Above/below-average rule. `aboveAverage: false` → below-average; provide `stdDev` for ±N stddev. */
export const addAverageRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: {
    aboveAverage?: boolean;
    equalAverage?: boolean;
    stdDev?: number;
    dxfId?: number;
    priority?: number;
  },
): ConditionalFormattingRule => {
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: 'aboveAverage',
      priority: opts.priority ?? nextCfPriority(ws),
      formulas: [],
      ...(opts.aboveAverage !== undefined ? { aboveAverage: opts.aboveAverage } : {}),
      ...(opts.equalAverage !== undefined ? { equalAverage: opts.equalAverage } : {}),
      ...(opts.stdDev !== undefined ? { stdDev: opts.stdDev } : {}),
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
    }),
  );
};

/** Duplicate-values rule (each cell whose value appears more than once gets the dxf). */
export const addDuplicateValuesRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: { dxfId?: number; priority?: number; unique?: boolean } = {},
): ConditionalFormattingRule => {
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: opts.unique ? 'uniqueValues' : 'duplicateValues',
      priority: opts.priority ?? nextCfPriority(ws),
      formulas: [],
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
    }),
  );
};

/** Free-form formula rule — any `=ISNUMBER(A1)`-style boolean expression. */
export const addFormulaRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: { formula: string; dxfId?: number; stopIfTrue?: boolean; priority?: number },
): ConditionalFormattingRule => {
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: 'expression',
      priority: opts.priority ?? nextCfPriority(ws),
      formulas: [opts.formula],
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
      ...(opts.stopIfTrue !== undefined ? { stopIfTrue: opts.stopIfTrue } : {}),
    }),
  );
};

/** Text-based rule (containsText / notContains / beginsWith / endsWith). */
export const addTextRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: { operator: TextOperator; text: string; dxfId?: number; priority?: number },
): ConditionalFormattingRule => {
  // Excel uses different `type` tokens for the four text operators.
  const typeMap: Record<TextOperator, ConditionalFormattingRuleType> = {
    containsText: 'containsText',
    notContains: 'notContainsText',
    beginsWith: 'beginsWith',
    endsWith: 'endsWith',
  };
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: typeMap[opts.operator],
      priority: opts.priority ?? nextCfPriority(ws),
      operator: opts.operator,
      text: opts.text,
      formulas: [],
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
    }),
  );
};
