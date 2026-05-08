// Server-only: load typedoc's JSON dump and reshape it into the section
// tree consumed by /api/* pages. Cached at module level so we only do
// the (~3 MB) parse once per build.

import apiData from '../server/api-data.json';
import { renderSignature, renderType, summaryToText } from './format';
import { classify, SECTIONS } from './sections';
import type {
  ApiItem,
  ApiKind,
  ApiMember,
  ApiModule,
  ApiParameter,
  ApiSection,
  ApiSubgroup,
} from './types';

/* eslint-disable @typescript-eslint/no-explicit-any */
type Reflection = any;

const TYPEDOC_KIND = {
  Variable: 32,
  Function: 64,
  Class: 128,
  Interface: 256,
  TypeAlias: 2_097_152,
} as const;

function kindOf(reflection: Reflection): ApiKind | null {
  switch (reflection.kind) {
    case TYPEDOC_KIND.Variable:
      return 'variable';
    case TYPEDOC_KIND.Function:
      return 'function';
    case TYPEDOC_KIND.Class:
      return 'class';
    case TYPEDOC_KIND.Interface:
      return 'interface';
    case TYPEDOC_KIND.TypeAlias:
      return 'type';
    default:
      return null;
  }
}

function moduleNameOf(name: string): ApiModule | null {
  if (name === 'index' || name === 'streaming' || name === 'node') return name;
  return null;
}

function buildParameters(sig: Reflection | undefined): ApiParameter[] | undefined {
  if (!sig?.parameters?.length) return undefined;
  return sig.parameters.map((p: Reflection): ApiParameter => {
    const param: ApiParameter = {
      name: p.name,
      type: renderType(p.type),
      optional: Boolean(p.flags?.isOptional || p.defaultValue !== undefined),
    };
    if (p.defaultValue !== undefined) param.defaultValue = String(p.defaultValue);
    const desc = summaryToText(p.comment?.summary);
    if (desc) param.description = desc;
    return param;
  });
}

function buildMembers(reflection: Reflection): ApiMember[] | undefined {
  // Interfaces / type literals expose their fields via `children`.
  const children: Reflection[] = reflection.children ?? [];
  const props = children.filter(
    (c) => c.kind === 1024 || c.kind === 2048 || c.kind === 65_536,
  );
  if (!props.length) return undefined;
  return props.map((p): ApiMember => {
    const m: ApiMember = {
      name: p.name,
      type: renderType(p.type),
      optional: Boolean(p.flags?.isOptional),
    };
    const desc = summaryToText(p.comment?.summary);
    if (desc) m.description = desc;
    return m;
  });
}

function renderItemSignature(reflection: Reflection): string {
  const kind = kindOf(reflection);
  switch (kind) {
    case 'function': {
      const sig = reflection.signatures?.[0];
      if (!sig) return `function ${reflection.name}(): unknown`;
      return `function ${renderSignature(sig, { name: reflection.name })}`;
    }
    case 'class':
      return `class ${reflection.name}`;
    case 'interface': {
      const members = reflection.children ?? [];
      if (!members.length) return `interface ${reflection.name} {}`;
      const lines = members.map((m: Reflection) => {
        const opt = m.flags?.isOptional ? '?' : '';
        const ro = m.flags?.isReadonly ? 'readonly ' : '';
        return `  ${ro}${m.name}${opt}: ${renderType(m.type)};`;
      });
      return `interface ${reflection.name} {\n${lines.join('\n')}\n}`;
    }
    case 'type': {
      return `type ${reflection.name} = ${renderType(reflection.type)}`;
    }
    case 'variable': {
      const t = renderType(reflection.type);
      return `const ${reflection.name}: ${t}`;
    }
    default:
      return reflection.name;
  }
}

function buildItem(reflection: Reflection, module: ApiModule): ApiItem | null {
  const kind = kindOf(reflection);
  if (!kind) return null;
  const source = reflection.sources?.[0];
  if (!source) return null;
  const sectionId = classify({
    name: reflection.name,
    module,
    sourceFile: source.fileName,
  });

  const description =
    summaryToText(reflection.comment?.summary) ||
    summaryToText(reflection.signatures?.[0]?.comment?.summary);

  const item: ApiItem = {
    id: reflection.id,
    name: reflection.name,
    kind,
    module,
    sectionId,
    sourceFile: source.fileName,
    sourceLine: source.line,
    sourceUrl: source.url,
    description,
    signature: renderItemSignature(reflection),
  };

  if (kind === 'function') {
    const sig = reflection.signatures?.[0];
    if (sig) {
      const params = buildParameters(sig);
      if (params) item.parameters = params;
      item.returnType = renderType(sig.type);
      const returnTag = sig.comment?.blockTags?.find(
        (t: Reflection) => t.tag === '@returns' || t.tag === '@return',
      );
      const returnDesc = summaryToText(returnTag?.content);
      if (returnDesc) item.returnDescription = returnDesc;
    }
  }

  if (kind === 'interface') {
    const members = buildMembers(reflection);
    if (members) item.members = members;
  }

  if (kind === 'type') {
    // Type literal aliases get their members rendered as a table too.
    if (reflection.type?.type === 'reflection') {
      const members = buildMembers(reflection.type.declaration);
      if (members) item.members = members;
    }
  }

  return item;
}

let cached: ApiSection[] | null = null;

function loadRawTypedoc(): Reflection {
  return apiData as Reflection;
}

export function loadApiSections(): ApiSection[] {
  if (cached) return cached;

  const raw: Reflection = loadRawTypedoc();

  // Build empty section buckets in declared order.
  const buckets = new Map<string, ApiItem[]>();
  for (const s of SECTIONS) buckets.set(s.id, []);

  const seen = new Set<string>();

  for (const moduleReflection of raw.children ?? []) {
    const module = moduleNameOf(moduleReflection.name);
    if (!module) continue;
    for (const child of moduleReflection.children ?? []) {
      const item = buildItem(child, module);
      if (!item) continue;
      // Same export can appear under multiple modules (e.g. loadWorkbook is
      // re-exported from `node`). De-duplicate by name+kind, preferring the
      // first occurrence (which follows our entry-point order: index → node
      // → streaming).
      const dedupKey = `${item.kind}:${item.name}`;
      if (seen.has(dedupKey)) continue;
      seen.add(dedupKey);
      buckets.get(item.sectionId)?.push(item);
    }
  }

  // Sort items inside each section: classes / interfaces / types first
  // (most browsable), then functions, then variables. Alphabetical ties.
  const KIND_ORDER: Record<ApiKind, number> = {
    class: 0,
    interface: 1,
    type: 2,
    function: 3,
    variable: 4,
  };
  for (const items of buckets.values()) {
    items.sort((a, b) => {
      const k = KIND_ORDER[a.kind] - KIND_ORDER[b.kind];
      if (k !== 0) return k;
      return a.name.localeCompare(b.name);
    });
  }

  cached = SECTIONS.map((s) => {
    const items = buckets.get(s.id) ?? [];
    return {
      id: s.id,
      title: s.title,
      description: s.description,
      items,
      subgroups: buildSubgroups(items),
    };
  }).filter((s) => s.items.length > 0);

  return cached;
}

function fileToSubgroupLabel(file: string): string {
  // src/worksheet/protected-ranges.ts -> "Protected ranges"
  const base = file.split('/').pop() ?? file;
  const stem = base.replace(/\.ts$/, '');
  const words = stem.split(/[-_]/);
  return words.map((w, i) => (i === 0 ? w[0]?.toUpperCase() + w.slice(1) : w)).join(' ');
}

function fileToSubgroupId(file: string): string {
  return file.replace(/^src\//, '').replace(/\.ts$/, '').replace(/\//g, '-');
}

function buildSubgroups(items: ApiItem[]): ApiSubgroup[] {
  if (items.length === 0) return [];
  const groups = new Map<string, ApiSubgroup>();
  for (const item of items) {
    const file = item.sourceFile;
    let g = groups.get(file);
    if (!g) {
      g = {
        label: fileToSubgroupLabel(file),
        id: fileToSubgroupId(file),
        sourceFile: file,
        items: [],
      };
      groups.set(file, g);
    }
    g.items.push(item);
  }
  // Sort subgroups: largest first usually most relevant; ties alphabetical.
  return [...groups.values()].sort((a, b) => {
    if (b.items.length !== a.items.length) return b.items.length - a.items.length;
    return a.label.localeCompare(b.label);
  });
}

export function loadApiSection(sectionId: string): ApiSection | null {
  return loadApiSections().find((s) => s.id === sectionId) ?? null;
}
