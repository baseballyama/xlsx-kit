// Normalized public API shapes consumed by /api/* pages. These are
// independent of typedoc's wire format so the route code never has to
// touch raw reflections.

export type ApiKind = 'function' | 'class' | 'interface' | 'type' | 'variable';

export type ApiModule = 'index' | 'streaming' | 'node';

export type ApiParameter = {
  name: string;
  type: string;
  optional: boolean;
  defaultValue?: string;
  description?: string;
};

export type ApiMember = {
  name: string;
  type: string;
  optional: boolean;
  description?: string;
};

export type ApiItem = {
  /** Stable typedoc id, used to link {@link} references. */
  id: number;
  name: string;
  kind: ApiKind;
  module: ApiModule;
  sectionId: string;
  /** Repo path like `src/public/load.ts`. */
  sourceFile: string;
  sourceLine: number;
  sourceUrl: string;
  /** Markdown summary (with [link](#anchor) substitutions where possible). */
  description: string;
  /** Single-line or multi-line TS-ish signature, ready for syntax highlighting. */
  signature: string;
  parameters?: ApiParameter[];
  returnType?: string;
  returnDescription?: string;
  /** For interface / type-literal kinds: enumerated property list. */
  members?: ApiMember[];
};

export type ApiSubgroup = {
  /** Display label for the H2, derived from the source file. */
  label: string;
  /** Stable id (used as anchor target). */
  id: string;
  /** Source file the items came from (relative to repo root). */
  sourceFile: string;
  items: ApiItem[];
};

export type ApiSection = {
  id: string;
  title: string;
  description: string;
  /** Flat item list, kept for fast counts. */
  items: ApiItem[];
  /** Items grouped by source file for the section page. */
  subgroups: ApiSubgroup[];
};

export type ApiSectionSummary = {
  id: string;
  title: string;
  description: string;
  itemCount: number;
};
