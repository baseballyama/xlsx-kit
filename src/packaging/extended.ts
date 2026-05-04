// `docProps/app.xml` — Office "extended" document properties. Per
// docs/plan/03-foundations.md §6.3.
//
// Mirrors openpyxl/openpyxl/packaging/extended.py. Most fields are flat
// (text or numeric). The four vector-shaped children — HeadingPairs,
// TitlesOfParts, HLinks, DigSig — are kept as raw XmlNode subtrees and
// round-tripped verbatim; modelling vt:vector explicitly is a phase-3
// concern at the earliest. See docs/plan/02-mapping.md row for
// `packaging/extended.py`.

import { defineSchema, type Schema } from '../schema/core';
import { fromTree, toTree } from '../schema/serialize';
import { XPROPS_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { serializeXml } from '../xml/serializer';
import type { XmlNode } from '../xml/tree';

export interface ExtendedProperties {
  application?: string;
  appVersion?: string;
  characters?: number;
  charactersWithSpaces?: number;
  company?: string;
  digitalSignature?: XmlNode;
  docSecurity?: number;
  hLinks?: XmlNode;
  headingPairs?: XmlNode;
  hiddenSlides?: number;
  hyperlinkBase?: string;
  hyperlinks?: XmlNode;
  hyperlinksChanged?: boolean;
  lines?: number;
  linksUpToDate?: boolean;
  manager?: string;
  mmClips?: number;
  notes?: number;
  pages?: number;
  paragraphs?: number;
  presentationFormat?: string;
  scaleCrop?: boolean;
  sharedDoc?: boolean;
  slides?: number;
  template?: string;
  titlesOfParts?: XmlNode;
  totalTime?: number;
  words?: number;
}

const ExtendedSchema: Schema<ExtendedProperties> = defineSchema<ExtendedProperties>({
  tagname: 'Properties',
  xmlNs: XPROPS_NS,
  attrs: {},
  // Element order chosen to match the Office output convention.
  elements: [
    { kind: 'text', key: 'template', xmlNs: XPROPS_NS, name: 'Template', primitive: 'string', optional: true },
    { kind: 'text', key: 'manager', xmlNs: XPROPS_NS, name: 'Manager', primitive: 'string', optional: true },
    { kind: 'text', key: 'company', xmlNs: XPROPS_NS, name: 'Company', primitive: 'string', optional: true },
    { kind: 'text', key: 'pages', xmlNs: XPROPS_NS, name: 'Pages', primitive: 'int', optional: true },
    { kind: 'text', key: 'words', xmlNs: XPROPS_NS, name: 'Words', primitive: 'int', optional: true },
    { kind: 'text', key: 'characters', xmlNs: XPROPS_NS, name: 'Characters', primitive: 'int', optional: true },
    {
      kind: 'text',
      key: 'presentationFormat',
      xmlNs: XPROPS_NS,
      name: 'PresentationFormat',
      primitive: 'string',
      optional: true,
    },
    { kind: 'text', key: 'lines', xmlNs: XPROPS_NS, name: 'Lines', primitive: 'int', optional: true },
    { kind: 'text', key: 'paragraphs', xmlNs: XPROPS_NS, name: 'Paragraphs', primitive: 'int', optional: true },
    { kind: 'text', key: 'slides', xmlNs: XPROPS_NS, name: 'Slides', primitive: 'int', optional: true },
    { kind: 'text', key: 'notes', xmlNs: XPROPS_NS, name: 'Notes', primitive: 'int', optional: true },
    { kind: 'text', key: 'totalTime', xmlNs: XPROPS_NS, name: 'TotalTime', primitive: 'int', optional: true },
    { kind: 'text', key: 'hiddenSlides', xmlNs: XPROPS_NS, name: 'HiddenSlides', primitive: 'int', optional: true },
    { kind: 'text', key: 'mmClips', xmlNs: XPROPS_NS, name: 'MMClips', primitive: 'int', optional: true },
    { kind: 'text', key: 'scaleCrop', xmlNs: XPROPS_NS, name: 'ScaleCrop', primitive: 'bool', optional: true },
    { kind: 'raw', key: 'headingPairs', xmlNs: XPROPS_NS, name: 'HeadingPairs', optional: true },
    { kind: 'raw', key: 'titlesOfParts', xmlNs: XPROPS_NS, name: 'TitlesOfParts', optional: true },
    { kind: 'raw', key: 'hLinks', xmlNs: XPROPS_NS, name: 'HLinks', optional: true },
    { kind: 'raw', key: 'hyperlinks', xmlNs: XPROPS_NS, name: 'Hyperlinks', optional: true },
    { kind: 'text', key: 'linksUpToDate', xmlNs: XPROPS_NS, name: 'LinksUpToDate', primitive: 'bool', optional: true },
    { kind: 'text', key: 'sharedDoc', xmlNs: XPROPS_NS, name: 'SharedDoc', primitive: 'bool', optional: true },
    {
      kind: 'text',
      key: 'hyperlinkBase',
      xmlNs: XPROPS_NS,
      name: 'HyperlinkBase',
      primitive: 'string',
      optional: true,
    },
    {
      kind: 'text',
      key: 'hyperlinksChanged',
      xmlNs: XPROPS_NS,
      name: 'HyperlinksChanged',
      primitive: 'bool',
      optional: true,
    },
    { kind: 'raw', key: 'digitalSignature', xmlNs: XPROPS_NS, name: 'DigSig', optional: true },
    { kind: 'text', key: 'application', xmlNs: XPROPS_NS, name: 'Application', primitive: 'string', optional: true },
    { kind: 'text', key: 'appVersion', xmlNs: XPROPS_NS, name: 'AppVersion', primitive: 'string', optional: true },
    { kind: 'text', key: 'docSecurity', xmlNs: XPROPS_NS, name: 'DocSecurity', primitive: 'int', optional: true },
    {
      kind: 'text',
      key: 'charactersWithSpaces',
      xmlNs: XPROPS_NS,
      name: 'CharactersWithSpaces',
      primitive: 'int',
      optional: true,
    },
  ],
});

export function makeExtendedProperties(): ExtendedProperties {
  return {};
}

export function extendedPropsToBytes(p: ExtendedProperties): Uint8Array {
  return serializeXml(toTree(p, ExtendedSchema));
}

export function extendedPropsFromBytes(bytes: Uint8Array | string): ExtendedProperties {
  return fromTree(parseXml(bytes), ExtendedSchema);
}
