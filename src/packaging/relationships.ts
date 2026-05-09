// `*.rels` files — Open Packaging Conventions relationship lists.
//
// Mirrors openpyxl/openpyxl/packaging/relationship.py. Relationship Ids
// (`rId{N}`) auto-increment using the smallest unused integer so the list stays
// deterministic even when callers mix in pre-assigned ids.

import { defineSchema, type Schema } from '../schema/core';
import { fromTree, toTree } from '../schema/serialize';
import { PKG_REL_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { serializeXml } from '../xml/serializer';

export interface Relationship {
  /** rIdN identifier; unique within the parent .rels file. */
  id: string;
  /** Full relationship type URI. */
  type: string;
  /** Target — relative part path or absolute URL when targetMode === 'External'. */
  target: string;
  /** Defaults to internal. */
  targetMode?: 'External' | 'Internal';
}

export interface Relationships {
  rels: Relationship[];
}

const RelationshipSchema = defineSchema<Relationship>({
  tagname: 'Relationship',
  xmlNs: PKG_REL_NS,
  attrs: {
    id: { kind: 'string', xmlName: 'Id' },
    type: { kind: 'string', xmlName: 'Type' },
    target: { kind: 'string', xmlName: 'Target' },
    targetMode: {
      kind: 'enum',
      values: ['External', 'Internal'],
      xmlName: 'TargetMode',
      optional: true,
    },
  },
  elements: [],
});

const RelationshipsSchema: Schema<Relationships> = defineSchema<Relationships>({
  tagname: 'Relationships',
  xmlNs: PKG_REL_NS,
  attrs: {},
  elements: [
    {
      kind: 'sequence',
      key: 'rels',
      itemName: 'Relationship',
      itemNs: PKG_REL_NS,
      itemSchema: () => RelationshipSchema,
    },
  ],
});

export function makeRelationships(): Relationships {
  return { rels: [] };
}

const RID_RE = /^rId(\d+)$/;

const allocateNextId = (rels: Relationships): string => {
  // Smallest positive integer not yet used as `rId{N}`. Linear scan is fine
  // because rels lists are small (tens of entries in real workbooks).
  const used = new Set<number>();
  for (const r of rels.rels) {
    const m = RID_RE.exec(r.id);
    if (m !== null) {
      const n = Number.parseInt(m[1] ?? '0', 10);
      if (Number.isFinite(n) && n > 0) used.add(n);
    }
  }
  let n = 1;
  while (used.has(n)) n++;
  return `rId${n}`;
};

/**
 * Append a relationship and return the resulting object. Auto-assigns the next
 * free `rId{N}` Id.
 */
export function appendRel(rels: Relationships, type: string, target: string, targetMode?: 'External'): Relationship {
  const next: Relationship = { id: allocateNextId(rels), type, target };
  if (targetMode !== undefined) next.targetMode = targetMode;
  rels.rels.push(next);
  return next;
}

export function findById(rels: Relationships, id: string): Relationship | undefined {
  for (const r of rels.rels) if (r.id === id) return r;
  return undefined;
}

export function findByType(rels: Relationships, type: string): Relationship | undefined {
  for (const r of rels.rels) if (r.type === type) return r;
  return undefined;
}

export function findAllByType(rels: Relationships, type: string): Relationship[] {
  const out: Relationship[] = [];
  for (const r of rels.rels) if (r.type === type) out.push(r);
  return out;
}

export function relsToBytes(rels: Relationships): Uint8Array {
  return serializeXml(toTree(rels, RelationshipsSchema));
}

export function relsFromBytes(bytes: Uint8Array | string): Relationships {
  return fromTree(parseXml(bytes), RelationshipsSchema);
}
