// `docProps/core.xml` — Dublin Core / Office core document properties.
// Per docs/plan/03-foundations.md §6.3.
//
// Each property lives in its own namespace (cp / dc / dcterms). The
// dcterms timestamps carry an `xsi:type="dcterms:W3CDTF"` marker which
// we reproduce as a fixed attr on those text elements via the schema
// layer's `text` element kind.

import { defineSchema, type Schema } from '../schema/core';
import { fromTree, toTree } from '../schema/serialize';
import { COREPROPS_NS, DCORE_NS, DCTERMS_NS, XSI_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { serializeXml } from '../xml/serializer';

/**
 * Set of properties exposed under `docProps/core.xml`. All fields are
 * optional; the workbook only emits those that are set. Timestamps are
 * stored as ISO-8601 strings (the W3CDTF subset) — no Date conversion
 * happens at this layer; phase-3 saveWorkbook is responsible for
 * stamping `modified` to `now()` on each save.
 */
export interface CoreProperties {
  category?: string;
  contentStatus?: string;
  /** ISO-8601 W3CDTF; auto-stamped on save in phase 3. */
  created?: string;
  creator?: string;
  description?: string;
  identifier?: string;
  keywords?: string;
  language?: string;
  lastModifiedBy?: string;
  /** ISO-8601 W3CDTF. */
  lastPrinted?: string;
  /** ISO-8601 W3CDTF; auto-stamped on save in phase 3. */
  modified?: string;
  revision?: string;
  subject?: string;
  title?: string;
  version?: string;
}

const W3CDTF_ATTRS: Record<string, string> = {
  [`{${XSI_NS}}type`]: 'dcterms:W3CDTF',
};

const CorePropertiesSchema: Schema<CoreProperties> = defineSchema<CoreProperties>({
  tagname: 'coreProperties',
  xmlNs: COREPROPS_NS,
  attrs: {},
  elements: [
    // openpyxl's child-namespace assignment, mirrored exactly:
    { kind: 'text', key: 'category', xmlNs: COREPROPS_NS, primitive: 'string', optional: true },
    { kind: 'text', key: 'contentStatus', xmlNs: COREPROPS_NS, primitive: 'string', optional: true },
    { kind: 'text', key: 'keywords', xmlNs: COREPROPS_NS, primitive: 'string', optional: true },
    { kind: 'text', key: 'lastModifiedBy', xmlNs: COREPROPS_NS, primitive: 'string', optional: true },
    {
      kind: 'text',
      key: 'lastPrinted',
      xmlNs: DCTERMS_NS,
      primitive: 'string',
      optional: true,
      attrs: W3CDTF_ATTRS,
    },
    { kind: 'text', key: 'revision', xmlNs: COREPROPS_NS, primitive: 'string', optional: true },
    { kind: 'text', key: 'version', xmlNs: COREPROPS_NS, primitive: 'string', optional: true },
    { kind: 'text', key: 'description', xmlNs: DCORE_NS, primitive: 'string', optional: true },
    { kind: 'text', key: 'identifier', xmlNs: DCORE_NS, primitive: 'string', optional: true },
    { kind: 'text', key: 'language', xmlNs: DCORE_NS, primitive: 'string', optional: true },
    { kind: 'text', key: 'subject', xmlNs: DCORE_NS, primitive: 'string', optional: true },
    { kind: 'text', key: 'title', xmlNs: DCORE_NS, primitive: 'string', optional: true },
    { kind: 'text', key: 'creator', xmlNs: DCORE_NS, primitive: 'string', optional: true },
    {
      kind: 'text',
      key: 'created',
      xmlNs: DCTERMS_NS,
      primitive: 'string',
      optional: true,
      attrs: W3CDTF_ATTRS,
    },
    {
      kind: 'text',
      key: 'modified',
      xmlNs: DCTERMS_NS,
      primitive: 'string',
      optional: true,
      attrs: W3CDTF_ATTRS,
    },
  ],
});

export function makeCoreProperties(): CoreProperties {
  return {};
}

export function corePropsToBytes(p: CoreProperties): Uint8Array {
  return serializeXml(toTree(p, CorePropertiesSchema));
}

export function corePropsFromBytes(bytes: Uint8Array | string): CoreProperties {
  return fromTree(parseXml(bytes), CorePropertiesSchema);
}

// ---- Workbook ergonomic helpers ----------------------------------------

import type { Workbook } from '../xlsx/workbook/workbook';

const ensureCoreProperties = (wb: Workbook): CoreProperties => {
  if (!wb.properties) wb.properties = {};
  return wb.properties;
};

/** Set the document author (Excel "File → Properties → Author"). */
export const setWorkbookCreator = (wb: Workbook, creator: string): void => {
  ensureCoreProperties(wb).creator = creator;
};

/** Set the document title. */
export const setWorkbookTitle = (wb: Workbook, title: string): void => {
  ensureCoreProperties(wb).title = title;
};

/** Set the document subject. */
export const setWorkbookSubject = (wb: Workbook, subject: string): void => {
  ensureCoreProperties(wb).subject = subject;
};

/** Set the document description / abstract. */
export const setWorkbookDescription = (wb: Workbook, description: string): void => {
  ensureCoreProperties(wb).description = description;
};

/** Set comma- or semicolon-separated keywords. */
export const setWorkbookKeywords = (wb: Workbook, keywords: string): void => {
  ensureCoreProperties(wb).keywords = keywords;
};

/** Set the "last modified by" name (defaults to creator if absent). */
export const setWorkbookLastModifiedBy = (wb: Workbook, name: string): void => {
  ensureCoreProperties(wb).lastModifiedBy = name;
};

/** Set the document category (e.g. "Reports", "Drafts"). */
export const setWorkbookCategory = (wb: Workbook, category: string): void => {
  ensureCoreProperties(wb).category = category;
};
