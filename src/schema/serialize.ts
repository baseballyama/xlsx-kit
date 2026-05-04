// Schema-driven (de)serialisation. One switch-on-kind walk in each
// direction; the schema does the work, no class-based descriptors.
//
// Per docs/plan/03-foundations.md §4 / docs/plan/01-architecture.md §5.3
// `toTree` and `fromTree` are pure functions over plain data. They run
// the same on the schema layer's own types and on user types in the
// styles / chart / drawing modules.

import { OpenXmlSchemaError } from '../utils/exceptions';
import { qname } from '../xml/namespaces';
import { el, type XmlNode } from '../xml/tree';
import type { AttrDef, ElementDef, Primitive, Schema } from './core';

// ---- coercion helpers -------------------------------------------------------

const coerceFromString = (raw: string, kind: Primitive | 'enum', def: AttrDef): unknown => {
  switch (kind) {
    case 'string':
      return raw;
    case 'int': {
      const n = Number.parseInt(raw, 10);
      if (!Number.isFinite(n)) throw new OpenXmlSchemaError(`expected an integer, got "${raw}"`);
      if (def.min !== undefined && n < def.min) throw new OpenXmlSchemaError(`int "${n}" is below min ${def.min}`);
      if (def.max !== undefined && n > def.max) throw new OpenXmlSchemaError(`int "${n}" is above max ${def.max}`);
      return n;
    }
    case 'float': {
      const n = Number.parseFloat(raw);
      if (!Number.isFinite(n)) throw new OpenXmlSchemaError(`expected a number, got "${raw}"`);
      if (def.min !== undefined && n < def.min) throw new OpenXmlSchemaError(`float "${n}" is below min ${def.min}`);
      if (def.max !== undefined && n > def.max) throw new OpenXmlSchemaError(`float "${n}" is above max ${def.max}`);
      return n;
    }
    case 'bool': {
      // openpyxl accepts 'true'/'t'/'1' or 'false'/'f'/'0' (case-insensitive).
      const v = raw.toLowerCase();
      if (v === 'true' || v === 't' || v === '1') return true;
      if (v === 'false' || v === 'f' || v === '0') return false;
      throw new OpenXmlSchemaError(`expected a boolean, got "${raw}"`);
    }
    case 'enum': {
      if (!def.values?.includes(raw)) {
        throw new OpenXmlSchemaError(`expected one of [${def.values?.join(', ') ?? ''}], got "${raw}"`);
      }
      return raw;
    }
  }
};

const coerceToString = (value: unknown, kind: Primitive | 'enum', def: AttrDef): string => {
  switch (kind) {
    case 'string':
      if (typeof value !== 'string') throw new OpenXmlSchemaError(`expected a string, got ${typeof value}`);
      return value;
    case 'int': {
      if (typeof value !== 'number' || !Number.isFinite(value) || !Number.isInteger(value)) {
        throw new OpenXmlSchemaError(`expected an integer, got ${String(value)}`);
      }
      return String(value);
    }
    case 'float': {
      if (typeof value !== 'number' || !Number.isFinite(value)) {
        throw new OpenXmlSchemaError(`expected a number, got ${String(value)}`);
      }
      return String(value);
    }
    case 'bool': {
      if (typeof value !== 'boolean') throw new OpenXmlSchemaError(`expected a boolean, got ${typeof value}`);
      // OOXML convention: '1' / '0'.
      return value ? '1' : '0';
    }
    case 'enum': {
      if (typeof value !== 'string' || !def.values || !def.values.includes(value)) {
        throw new OpenXmlSchemaError(`enum value out of range: ${String(value)}`);
      }
      return value;
    }
  }
};

const attrXmlKey = (key: string, def: AttrDef): string => qname(def.xmlNs ?? '', def.xmlName ?? key);
const elementXmlName = (def: ElementDef, fallbackNs: string | undefined): string => {
  const local = def.kind === 'sequence' ? (def.container?.name ?? def.itemName) : (def.name ?? def.key);
  const ns = def.kind === 'sequence' ? (def.container?.xmlNs ?? def.itemNs ?? fallbackNs) : (def.xmlNs ?? fallbackNs);
  return qname(ns ?? '', local);
};

// ---- toTree -----------------------------------------------------------------

/** Serialise a value of `T` into an {@link XmlNode} per the supplied schema. */
export function toTree<T>(value: T, schema: Schema<T>): XmlNode {
  const v = (schema.preSerialize ? schema.preSerialize(value) : value) as Record<string, unknown>;
  const node = el(qname(schema.xmlNs ?? '', schema.tagname));

  for (const [key, def] of Object.entries(schema.attrs) as Array<[string, AttrDef | undefined]>) {
    if (def === undefined) continue;
    const raw = v[key];
    if (raw === undefined) {
      if (def.optional) continue;
      throw new OpenXmlSchemaError(`<${schema.tagname}>: required attribute "${key}" is missing`);
    }
    node.attrs[attrXmlKey(key, def)] = coerceToString(raw, def.kind, def);
  }

  for (const def of schema.elements) {
    const raw = v[def.key];
    if (raw === undefined) {
      // 'empty' has no `optional`; presence is binary on the value itself.
      // 'sequence' is implicitly optional (undefined / [] = no items).
      if ((def.kind === 'text' || def.kind === 'object') && !def.optional) {
        throw new OpenXmlSchemaError(`<${schema.tagname}>: required element "${def.key}" is missing`);
      }
      continue;
    }
    switch (def.kind) {
      case 'text': {
        const text = coerceToString(raw, def.primitive, { kind: def.primitive });
        node.children.push(el(elementXmlName(def, schema.xmlNs), def.attrs ?? {}, [], text));
        break;
      }
      case 'empty': {
        if (raw === true) node.children.push(el(elementXmlName(def, schema.xmlNs), {}, []));
        break;
      }
      case 'object': {
        const sub = toTree(raw as Record<string, unknown>, def.schema());
        // Allow the schema to declare its own tagname OR be addressed
        // by the parent under a different element name. We honour the
        // schema's tagname (it's the canonical identity for that type).
        if (def.name !== undefined && sub.name !== qname(def.xmlNs ?? schema.xmlNs ?? '', def.name)) {
          // Re-tag the produced element to the parent-defined name.
          sub.name = qname(def.xmlNs ?? schema.xmlNs ?? '', def.name);
        } else if (def.name === undefined && def.key !== sub.name.replace(/^\{[^}]*\}/, '')) {
          // No explicit `name` — use the property key as the local name.
          sub.name = qname(def.xmlNs ?? schema.xmlNs ?? '', def.key);
        }
        node.children.push(sub);
        break;
      }
      case 'sequence': {
        const items = raw as ReadonlyArray<Record<string, unknown>>;
        const itemSchema = def.itemSchema();
        const built = items.map((item) => {
          const sub = toTree(item, itemSchema);
          // Re-tag to the declared itemName / itemNs.
          sub.name = qname(def.itemNs ?? itemSchema.xmlNs ?? '', def.itemName);
          return sub;
        });
        if (def.container !== undefined) {
          const containerAttrs: Record<string, string> = {};
          if (def.container.count) containerAttrs['count'] = String(items.length);
          node.children.push(el(elementXmlName(def, schema.xmlNs), containerAttrs, built));
        } else {
          for (const c of built) node.children.push(c);
        }
        break;
      }
    }
  }

  return node;
}

// ---- fromTree ---------------------------------------------------------------

const childByName = (node: XmlNode, fullName: string): XmlNode | undefined => {
  for (const c of node.children) if (c.name === fullName) return c;
  return undefined;
};

const childrenByName = (node: XmlNode, fullName: string): XmlNode[] => {
  const out: XmlNode[] = [];
  for (const c of node.children) if (c.name === fullName) out.push(c);
  return out;
};

/** Materialise a value of `T` from an {@link XmlNode} per the supplied schema. */
export function fromTree<T>(node: XmlNode, schema: Schema<T>): T {
  const out: Record<string, unknown> = {};
  const expectedTag = qname(schema.xmlNs ?? '', schema.tagname);
  if (node.name !== expectedTag) {
    throw new OpenXmlSchemaError(`fromTree: expected <${schema.tagname}> (Clark "${expectedTag}"), got "${node.name}"`);
  }

  for (const [key, def] of Object.entries(schema.attrs) as Array<[string, AttrDef | undefined]>) {
    if (def === undefined) continue;
    const xmlKey = attrXmlKey(key, def);
    const raw = node.attrs[xmlKey];
    if (raw === undefined) {
      if (def.default !== undefined) out[key] = def.default;
      else if (!def.optional) {
        throw new OpenXmlSchemaError(`<${schema.tagname}>: required attribute "${key}" is missing`);
      }
      continue;
    }
    out[key] = coerceFromString(raw, def.kind, def);
  }

  for (const def of schema.elements) {
    const fullName = elementXmlName(def, schema.xmlNs);
    switch (def.kind) {
      case 'text': {
        const child = childByName(node, fullName);
        if (child === undefined) {
          if (def.default !== undefined) out[def.key] = def.default;
          else if (!def.optional) {
            throw new OpenXmlSchemaError(`<${schema.tagname}>: required text element "${def.key}" is missing`);
          }
          break;
        }
        out[def.key] = coerceFromString(child.text ?? '', def.primitive, { kind: def.primitive });
        break;
      }
      case 'empty': {
        const child = childByName(node, fullName);
        out[def.key] = child !== undefined;
        break;
      }
      case 'object': {
        const child = childByName(node, fullName);
        if (child === undefined) {
          if (!def.optional) {
            throw new OpenXmlSchemaError(`<${schema.tagname}>: required element "${def.key}" is missing`);
          }
          break;
        }
        const itemSchema = def.schema();
        // Address the child by the parent-declared name even if the
        // child schema names itself differently.
        const renamed: XmlNode =
          child.name === qname(itemSchema.xmlNs ?? '', itemSchema.tagname)
            ? child
            : { ...child, name: qname(itemSchema.xmlNs ?? '', itemSchema.tagname) };
        out[def.key] = fromTree(renamed, itemSchema);
        break;
      }
      case 'sequence': {
        const itemSchema = def.itemSchema();
        const itemFullName = qname(def.itemNs ?? itemSchema.xmlNs ?? '', def.itemName);
        const itemSchemaName = qname(itemSchema.xmlNs ?? '', itemSchema.tagname);
        const containerNode = def.container !== undefined ? childByName(node, fullName) : node;
        const items: unknown[] = [];
        if (containerNode !== undefined) {
          const matches = childrenByName(containerNode, itemFullName);
          for (const m of matches) {
            const renamed: XmlNode = m.name === itemSchemaName ? m : { ...m, name: itemSchemaName };
            items.push(fromTree(renamed, itemSchema));
          }
        }
        out[def.key] = items;
        break;
      }
    }
  }

  const built = out as unknown as T;
  return schema.postParse ? schema.postParse(built, node) : built;
}
