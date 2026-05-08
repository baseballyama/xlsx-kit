// Workbook-level <functionGroups>. Per ECMA-376 §18.2.14.
//
// Excel registers built-in function groups (Math/Statistical/Logical/...)
// implicitly with a `builtInGroupCount` count and supports user-defined
// XLL function groups appended after that count. Most workbooks don't
// carry this element at all.

export interface FunctionGroup {
  name: string;
}

export interface FunctionGroups {
  /** Number of built-in groups Excel reserves before user entries (default 16). */
  builtInGroupCount?: number;
  groups: FunctionGroup[];
}

export const makeFunctionGroup = (name: string): FunctionGroup => ({ name });

export const makeFunctionGroups = (opts: Partial<FunctionGroups> = {}): FunctionGroups => ({
  groups: opts.groups?.slice() ?? [],
  ...(opts.builtInGroupCount !== undefined ? { builtInGroupCount: opts.builtInGroupCount } : {}),
});
