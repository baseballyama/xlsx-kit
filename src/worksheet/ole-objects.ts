// Worksheet-level <oleObjects> + <controls>. Per ECMA-376 §18.3.1.61
// and §18.3.1.27.
//
// Both elements are lists of object references where each entry has a
// `shapeId` + `r:id` plus a child `<objectPr>` / `<controlPr>` element
// that holds the anchor / display properties. The objectPr/controlPr
// children pull in the spreadsheet-drawing namespace, so we round-trip
// them as opaque XmlNode payloads rather than fully modeling the
// nested anchor schema. The top-level attrs are typed for editor access.

import type { XmlNode } from '../xml/tree';

export type OleDvAspect = 'DVASPECT_CONTENT' | 'DVASPECT_ICON';
export type OleUpdateMode = 'OLEUPDATE_ALWAYS' | 'OLEUPDATE_ONCALL';

export interface OleObject {
  /** Unique shape id assigned by Excel — required. */
  shapeId: number;
  /** rels link to the embedded OLE blob. */
  rId?: string;
  progId?: string;
  dvAspect?: OleDvAspect;
  link?: string;
  oleUpdate?: OleUpdateMode;
  autoLoad?: boolean;
  /**
   * Optional `<objectPr>` child preserved verbatim. Modeling its
   * `<anchor>` schema in detail is deferred — this preserves the
   * round-trip without re-deriving the anchor attrs.
   */
  objectPr?: XmlNode;
}

export interface FormControl {
  shapeId: number;
  rId?: string;
  /** ECMA-376 §18.3.1.27 — name shown in the form-control name box. */
  name?: string;
  /**
   * Optional `<controlPr>` child preserved verbatim (similar to
   * `objectPr` for OLE objects).
   */
  controlPr?: XmlNode;
}

export const makeOleObject = (opts: Partial<OleObject> & { shapeId: number }): OleObject => ({
  shapeId: opts.shapeId,
  ...(opts.rId !== undefined ? { rId: opts.rId } : {}),
  ...(opts.progId !== undefined ? { progId: opts.progId } : {}),
  ...(opts.dvAspect !== undefined ? { dvAspect: opts.dvAspect } : {}),
  ...(opts.link !== undefined ? { link: opts.link } : {}),
  ...(opts.oleUpdate !== undefined ? { oleUpdate: opts.oleUpdate } : {}),
  ...(opts.autoLoad !== undefined ? { autoLoad: opts.autoLoad } : {}),
  ...(opts.objectPr !== undefined ? { objectPr: opts.objectPr } : {}),
});

export const makeFormControl = (opts: Partial<FormControl> & { shapeId: number }): FormControl => ({
  shapeId: opts.shapeId,
  ...(opts.rId !== undefined ? { rId: opts.rId } : {}),
  ...(opts.name !== undefined ? { name: opts.name } : {}),
  ...(opts.controlPr !== undefined ? { controlPr: opts.controlPr } : {}),
});
