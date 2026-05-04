// Lightweight XML tree representation used between the parser, the
// schema layer, and the serializer.
//
// Per docs/plan/03-foundations.md §3.2: plain object, no DOM, names in
// Clark notation `{ns}local` so namespace-aware code does not depend on
// runtime prefix bookkeeping.

import { qname } from './namespaces';

export interface XmlNode {
  /** Clark-notation qualified name: `{namespace}local` or just `local`. */
  name: string;
  /** Attribute table; values are always strings (no value coercion at this layer). */
  attrs: Record<string, string>;
  /**
   * Optional element text. When an element has both text and child
   * elements, callers should set `text` and rely on `children` for
   * mixed content; the serializer emits `text` then children.
   */
  text?: string;
  /** Child element nodes in document order. */
  children: XmlNode[];
}

/**
 * Build an XmlNode from primitive bits. Attributes / children default
 * to empty; pass `undefined` for `text` to omit the text node.
 *
 * The element's name is always supplied in Clark notation. Use the
 * {@link qname} helper from `./namespaces` to keep call sites readable.
 */
export function el(
  name: string,
  attrs: Readonly<Record<string, string | number | boolean | null | undefined>> = {},
  children: ReadonlyArray<XmlNode> = [],
  text?: string,
): XmlNode {
  const out: XmlNode = {
    name,
    attrs: normaliseAttrs(attrs),
    children: children.slice(),
  };
  if (text !== undefined) out.text = text;
  return out;
}

const normaliseAttrs = (
  attrs: Readonly<Record<string, string | number | boolean | null | undefined>>,
): Record<string, string> => {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null) continue;
    out[k] = typeof v === 'string' ? v : String(v);
  }
  return out;
};

/**
 * Convenience: element with a Clark-notation name composed from a namespace
 * URI and local name.
 */
export function elNs(
  namespace: string | undefined,
  local: string,
  attrs?: Readonly<Record<string, string | number | boolean | null | undefined>>,
  children?: ReadonlyArray<XmlNode>,
  text?: string,
): XmlNode {
  return el(qname(namespace, local), attrs, children, text);
}

/** Locate the first child matching the supplied Clark-notation name. */
export function findChild(node: XmlNode, name: string): XmlNode | undefined {
  for (const c of node.children) if (c.name === name) return c;
  return undefined;
}

/** All children matching the supplied Clark-notation name, in document order. */
export function findChildren(node: XmlNode, name: string): XmlNode[] {
  const out: XmlNode[] = [];
  for (const c of node.children) if (c.name === name) out.push(c);
  return out;
}

/**
 * Append `child` to `parent`. Mutates and returns `parent` for chaining
 * during construction. Avoid in hot paths — for cell writing the worksheet
 * writer goes through a templated emitter rather than building XmlNode
 * trees per cell (docs/plan/01-architecture.md §7.2).
 */
export function appendChild(parent: XmlNode, child: XmlNode): XmlNode {
  parent.children.push(child);
  return parent;
}
