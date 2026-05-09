// Worksheet-level <customProperties> and <webPublishItems>.
//
// Both elements live near the bottom of <worksheet> (after tableParts, before
// extLst per ECMA-376 §18.3.1.43 / §18.3.1.97). Each is a thin shell over a
// list of children. customProperty references a Custom XML part via `r:id` —
// the underlying rel is already preserved by the worksheet's `relsExtras`
// machinery, so we just have to keep the inline element from leaking.

/**
 * One <customProperty>. The `rId` points at a Custom XML part registered in the
 * worksheet rels (e.g. for SharePoint sync metadata).
 */
export interface WorksheetCustomProperty {
  name: string;
  /** Worksheet-rels rId pointing at the Custom XML part backing this entry. */
  rId?: string;
}

export interface WebPublishItem {
  id: number;
  divId: string;
  sourceType: 'sheet' | 'printArea' | 'autoFilter' | 'range' | 'chart' | 'pivotTable' | 'query' | 'label';
  sourceRef?: string;
  sourceObject?: string;
  destinationFile: string;
  title?: string;
  autoRepublish?: boolean;
}

export const makeWorksheetCustomProperty = (
  opts: WorksheetCustomProperty,
): WorksheetCustomProperty => ({
  name: opts.name,
  ...(opts.rId !== undefined ? { rId: opts.rId } : {}),
});

export const makeWebPublishItem = (opts: WebPublishItem): WebPublishItem => ({
  id: opts.id,
  divId: opts.divId,
  sourceType: opts.sourceType,
  destinationFile: opts.destinationFile,
  ...(opts.sourceRef !== undefined ? { sourceRef: opts.sourceRef } : {}),
  ...(opts.sourceObject !== undefined ? { sourceObject: opts.sourceObject } : {}),
  ...(opts.title !== undefined ? { title: opts.title } : {}),
  ...(opts.autoRepublish !== undefined ? { autoRepublish: opts.autoRepublish } : {}),
});
