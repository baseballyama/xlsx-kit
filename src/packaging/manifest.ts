// [Content_Types].xml — the package manifest. Tracks default content
// types per file extension and per-part Override entries. Per
// docs/plan/03-foundations.md §6.1.

import { defineSchema, type Schema } from '../schema/core';
import { fromTree, toTree } from '../schema/serialize';
import { CONTYPES_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { serializeXml } from '../xml/serializer';

export interface DefaultEntry {
  /** File extension without leading dot (e.g. "rels", "xml", "png"). */
  ext: string;
  contentType: string;
}

export interface OverrideEntry {
  /** Part name path including leading slash (e.g. "/xl/workbook.xml"). */
  partName: string;
  contentType: string;
}

export interface Manifest {
  defaults: DefaultEntry[];
  overrides: OverrideEntry[];
}

const DefaultSchema = defineSchema<DefaultEntry>({
  tagname: 'Default',
  xmlNs: CONTYPES_NS,
  attrs: {
    ext: { kind: 'string', xmlName: 'Extension' },
    contentType: { kind: 'string', xmlName: 'ContentType' },
  },
  elements: [],
});

const OverrideSchema = defineSchema<OverrideEntry>({
  tagname: 'Override',
  xmlNs: CONTYPES_NS,
  attrs: {
    partName: { kind: 'string', xmlName: 'PartName' },
    contentType: { kind: 'string', xmlName: 'ContentType' },
  },
  elements: [],
});

const ManifestSchema: Schema<Manifest> = defineSchema<Manifest>({
  tagname: 'Types',
  xmlNs: CONTYPES_NS,
  attrs: {},
  elements: [
    {
      kind: 'sequence',
      key: 'defaults',
      itemName: 'Default',
      itemNs: CONTYPES_NS,
      itemSchema: () => DefaultSchema,
    },
    {
      kind: 'sequence',
      key: 'overrides',
      itemName: 'Override',
      itemNs: CONTYPES_NS,
      itemSchema: () => OverrideSchema,
    },
  ],
});

export function makeManifest(): Manifest {
  return { defaults: [], overrides: [] };
}

/** Register a default content type for an extension. Idempotent. */
export function addDefault(m: Manifest, ext: string, contentType: string): void {
  for (const d of m.defaults) {
    if (d.ext === ext) {
      if (d.contentType !== contentType) {
        d.contentType = contentType; // newer wins; matches openpyxl behaviour
      }
      return;
    }
  }
  m.defaults.push({ ext, contentType });
}

/** Register an Override for a specific part. Idempotent. */
export function addOverride(m: Manifest, partName: string, contentType: string): void {
  for (const o of m.overrides) {
    if (o.partName === partName) {
      if (o.contentType !== contentType) o.contentType = contentType;
      return;
    }
  }
  m.overrides.push({ partName, contentType });
}

export function findOverride(m: Manifest, partName: string): OverrideEntry | undefined {
  for (const o of m.overrides) if (o.partName === partName) return o;
  return undefined;
}

/** Find the first Override matching a content type. */
export function findOverrideByContentType(m: Manifest, contentType: string): OverrideEntry | undefined {
  for (const o of m.overrides) if (o.contentType === contentType) return o;
  return undefined;
}

export function manifestToBytes(m: Manifest): Uint8Array {
  return serializeXml(toTree(m, ManifestSchema));
}

export function manifestFromBytes(bytes: Uint8Array | string): Manifest {
  return fromTree(parseXml(bytes), ManifestSchema);
}
