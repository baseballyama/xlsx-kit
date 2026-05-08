// Worksheet-level <smartTags>. Per ECMA-376 §18.3.1.93.
//
// Per-cell smart-tag annotations from Excel 2003. The element is
// nested:
//   <smartTags>
//     <cellSmartTags r="A1">
//       <cellSmartTag type="0" deleted="0" xmlBased="0">
//         <cellSmartTagPr key="…" val="…"/>
//       </cellSmartTag>
//     </cellSmartTags>
//   </smartTags>
// Almost never seen in modern files; the workbook-level smartTagTypes
// list registers the schema, this element pins individual cells.

export interface CellSmartTagProperty {
  key: string;
  val: string;
}

export interface CellSmartTag {
  /** 0-based index into the workbook's smartTagTypes list. */
  type: number;
  properties: CellSmartTagProperty[];
  deleted?: boolean;
  xmlBased?: boolean;
}

export interface CellSmartTags {
  /** Single-cell ref ("A1"). */
  ref: string;
  tags: CellSmartTag[];
}

export const makeCellSmartTagProperty = (key: string, val: string): CellSmartTagProperty => ({ key, val });

export const makeCellSmartTag = (opts: Partial<CellSmartTag> & { type: number }): CellSmartTag => ({
  type: opts.type,
  properties: opts.properties?.slice() ?? [],
  ...(opts.deleted !== undefined ? { deleted: opts.deleted } : {}),
  ...(opts.xmlBased !== undefined ? { xmlBased: opts.xmlBased } : {}),
});

export const makeCellSmartTags = (opts: Partial<CellSmartTags> & { ref: string }): CellSmartTags => ({
  ref: opts.ref,
  tags: opts.tags?.slice() ?? [],
});
