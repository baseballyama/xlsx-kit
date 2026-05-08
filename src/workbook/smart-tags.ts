// Workbook-level <smartTagPr> + <smartTagTypes>. Per ECMA-376 §18.2.26
// / §18.2.27. Smart tags were Excel 2003's auto-recognized data
// (stock symbols, dates, names) and are deprecated, but the elements
// still appear in some legacy workbooks.

export type SmartTagShowMode = 'all' | 'noIndicator';

export interface SmartTagProperties {
  /** Embed smart tags into the workbook on save. */
  embed?: boolean;
  show?: SmartTagShowMode;
}

export interface SmartTagType {
  namespaceUri?: string;
  name?: string;
  url?: string;
}

export const makeSmartTagProperties = (opts: SmartTagProperties = {}): SmartTagProperties => ({ ...opts });

export const makeSmartTagType = (opts: SmartTagType = {}): SmartTagType => ({ ...opts });
