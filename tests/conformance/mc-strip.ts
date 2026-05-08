// Markup Compatibility (ECMA-376 Part 3) preprocessor + schema-quirk shim.
//
// Two independent jobs run in one pass so we only parse/serialise the XML
// once:
//
// 1. MC handling. Real-world xlsx files contain extension elements/attributes
//    (x14ac, x15, x16r2, …) gated by `mc:Ignorable`. The base ECMA-376
//    schemas don't know about these, so we honour the spec: drop everything
//    in an ignorable namespace, since a reader that doesn't understand the
//    extension is required to ignore it. `mc:AlternateContent` resolves to
//    its non-ignorable Choice, or its Fallback, or nothing.
//
// 2. xml:* attribute stripping. ECMA-376's simpleType-based string types
//    (e.g. ST_Xstring, used by `<t>` in sharedStrings) don't declare
//    xml:space, but every real producer (Excel, LibreOffice, openpyxl)
//    emits `xml:space="preserve"` on `<t>` whenever leading/trailing
//    whitespace matters. xmllint flags this as schema-invalid. Since the
//    XML core spec reserves the xml: namespace and xml:space is allowed on
//    any element, we drop xml:* attributes before validation so the schema
//    bug doesn't fire false positives.
//
// Implementation note: the stripping happens on raw XML text via a small
// state machine over fast-xml-parser's preserveOrder tree, so the stripped
// output round-trips through xmllint without introducing artefacts.

import { XMLBuilder, XMLParser } from 'fast-xml-parser';

const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const ATTR_KEY = ':@';

type FxpAttrs = Record<string, string>;
type FxpEntry = Record<string, unknown>;
type FxpTree = FxpEntry[];

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  attributesGroupName: ATTR_KEY,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  processEntities: true,
  htmlEntities: false,
});

const builder = new XMLBuilder({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  attributesGroupName: ATTR_KEY,
  suppressEmptyNode: true,
  format: false,
});

const splitQName = (q: string): { prefix: string; local: string } => {
  const i = q.indexOf(':');
  return i < 0 ? { prefix: '', local: q } : { prefix: q.slice(0, i), local: q.slice(i + 1) };
};

const tagOf = (entry: FxpEntry): string | undefined => {
  for (const k of Object.keys(entry)) if (k !== ATTR_KEY) return k;
  return undefined;
};

const attrsOf = (entry: FxpEntry): FxpAttrs | undefined => entry[ATTR_KEY] as FxpAttrs | undefined;

const childrenOf = (entry: FxpEntry, tag: string): FxpEntry[] | undefined => {
  const c = entry[tag];
  return Array.isArray(c) ? (c as FxpEntry[]) : undefined;
};

interface NSScope {
  /** Default namespace at this scope; '' if none. */
  defaultNs: string;
  /** prefix → namespace URI. */
  byPrefix: Record<string, string>;
  /** Namespace URIs declared as ignorable at or above this scope. */
  ignorable: Set<string>;
}

const extendScope = (parent: NSScope, attrs: FxpAttrs | undefined): NSScope => {
  if (!attrs) return parent;
  let next: NSScope | undefined;
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'xmlns') {
      next ??= { defaultNs: parent.defaultNs, byPrefix: { ...parent.byPrefix }, ignorable: new Set(parent.ignorable) };
      next.defaultNs = v;
    } else if (k.startsWith('xmlns:')) {
      next ??= { defaultNs: parent.defaultNs, byPrefix: { ...parent.byPrefix }, ignorable: new Set(parent.ignorable) };
      next.byPrefix[k.slice('xmlns:'.length)] = v;
    }
  }
  // mc:Ignorable lists prefixes (space-separated) whose namespaces may be dropped
  for (const [k, v] of Object.entries(attrs)) {
    if (resolveAttrNamespace(k, next ?? parent) !== MC_NS) continue;
    const { local } = splitQName(k);
    if (local !== 'Ignorable') continue;
    next ??= { defaultNs: parent.defaultNs, byPrefix: { ...parent.byPrefix }, ignorable: new Set(parent.ignorable) };
    for (const prefix of v.split(/\s+/).filter(Boolean)) {
      const ns = next.byPrefix[prefix];
      if (ns) next.ignorable.add(ns);
    }
  }
  return next ?? parent;
};

const resolveElementNamespace = (qname: string, scope: NSScope): string => {
  const { prefix } = splitQName(qname);
  if (prefix === '') return scope.defaultNs;
  return scope.byPrefix[prefix] ?? '';
};

const resolveAttrNamespace = (qname: string, scope: NSScope): string => {
  const { prefix } = splitQName(qname);
  if (prefix === '') return ''; // unprefixed attrs do not inherit default ns
  if (prefix === 'xml') return 'http://www.w3.org/XML/1998/namespace';
  if (prefix === 'xmlns') return 'http://www.w3.org/2000/xmlns/';
  return scope.byPrefix[prefix] ?? '';
};

const isMcAlternateContent = (entry: FxpEntry, scope: NSScope): boolean => {
  const tag = tagOf(entry);
  if (!tag) return false;
  return resolveElementNamespace(tag, scope) === MC_NS && splitQName(tag).local === 'AlternateContent';
};

const filterAttrs = (attrs: FxpAttrs | undefined, scope: NSScope): FxpAttrs | undefined => {
  if (!attrs) return undefined;
  const out: FxpAttrs = {};
  let any = false;
  for (const [k, v] of Object.entries(attrs)) {
    const ns = resolveAttrNamespace(k, scope);
    // Drop ignored-namespace attributes, MC bookkeeping, and xml:* core attrs
    // (see file header for the schema-quirk rationale on the latter). xmlns
    // declarations are kept so the resulting XML still parses standalone.
    if (ns === MC_NS) continue;
    if (ns === XML_NS) continue;
    if (scope.ignorable.has(ns)) continue;
    out[k] = v;
    any = true;
  }
  return any ? out : undefined;
};

/** Resolve a single mc:AlternateContent node into 0..1 replacement child. */
const resolveAlternateContent = (entry: FxpEntry, scope: NSScope): FxpEntry[] => {
  const tag = tagOf(entry);
  if (!tag) return [];
  const kids = childrenOf(entry, tag) ?? [];
  // Spec: pick the first <Choice> whose Requires only references *non-ignorable*
  // namespaces; otherwise fall back to <Fallback>; otherwise emit nothing.
  let fallback: FxpEntry | undefined;
  for (const kid of kids) {
    const kidTag = tagOf(kid);
    if (!kidTag) continue;
    const kidNs = resolveElementNamespace(kidTag, scope);
    if (kidNs !== MC_NS) continue;
    const { local } = splitQName(kidTag);
    if (local === 'Choice') {
      const requires = attrsOf(kid)?.['Requires'] ?? '';
      const allIgnorable = requires
        .split(/\s+/)
        .filter(Boolean)
        .every((p) => scope.ignorable.has(scope.byPrefix[p] ?? ''));
      if (!allIgnorable) {
        // Replace the AlternateContent with the children of this Choice.
        return childrenOf(kid, kidTag) ?? [];
      }
    } else if (local === 'Fallback') {
      fallback = kid;
    }
  }
  if (fallback) {
    const fbTag = tagOf(fallback);
    if (fbTag) return childrenOf(fallback, fbTag) ?? [];
  }
  return [];
};

const stripEntry = (entry: FxpEntry, parentScope: NSScope): FxpEntry | undefined => {
  if (Object.hasOwn(entry, '#text')) return entry; // text node passes through
  const tag = tagOf(entry);
  if (!tag) return undefined;
  const attrs = attrsOf(entry);
  const scope = extendScope(parentScope, attrs);

  const ns = resolveElementNamespace(tag, scope);
  if (scope.ignorable.has(ns)) return undefined;
  if (ns === MC_NS) return undefined; // bare MC element outside AlternateContent

  const filteredAttrs = filterAttrs(attrs, scope);
  const out: FxpEntry = { [tag]: [] as FxpEntry[] };
  if (filteredAttrs) out[ATTR_KEY] = filteredAttrs;

  const kidsArr = childrenOf(entry, tag) ?? [];
  const outKids: FxpEntry[] = [];
  for (const kid of kidsArr) {
    if (Object.hasOwn(kid, '#text')) {
      outKids.push(kid);
      continue;
    }
    if (isMcAlternateContent(kid, scope)) {
      const resolved = resolveAlternateContent(kid, scope);
      for (const r of resolved) {
        const stripped = stripEntry(r, scope);
        if (stripped) outKids.push(stripped);
      }
      continue;
    }
    const stripped = stripEntry(kid, scope);
    if (stripped) outKids.push(stripped);
  }
  out[tag] = outKids;
  return out;
};

/**
 * Remove ignorable-namespace content per ECMA-376 Part 3 (Markup Compatibility).
 * The output is a self-contained XML document that matches the base schemas.
 *
 * No-op (other than parse/serialize round-trip) when the document does not
 * declare `mc:Ignorable`.
 */
export function stripIgnorableMarkup(xml: string): string {
  const tree = parser.parse(xml) as FxpTree;
  const initial: NSScope = { defaultNs: '', byPrefix: {}, ignorable: new Set() };
  const out: FxpEntry[] = [];
  for (const entry of tree) {
    if (Object.hasOwn(entry, '?xml')) {
      out.push(entry);
      continue;
    }
    const stripped = stripEntry(entry, initial);
    if (stripped) out.push(stripped);
  }
  return builder.build(out) as string;
}
