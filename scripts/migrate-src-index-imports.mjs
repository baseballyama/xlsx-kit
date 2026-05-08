#!/usr/bin/env node
// One-shot codemod: rewrites `import { ... } from '<rel>/src/index'` to
// per-section subpath imports against the section indexes.
//
// Walks tests/ and rewrites each `import` statement that targets the
// removed barrel src/index.ts. Symbols are mapped to their owning
// section by parsing each src/<section>/index.ts at runtime.
//
// Usage: node scripts/migrate-src-index-imports.mjs

import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const ROOT = resolve(__dirname, '..');

const SECTIONS = [
  'cell',
  'chart',
  'chartsheet',
  'drawing',
  'io',
  'packaging',
  'schema',
  'streaming',
  'styles',
  'utils',
  'workbook',
  'worksheet',
  'xml',
  'zip',
];

// Build symbol -> section map by parsing each section's public index.ts.
function parseExports(text) {
  const symbols = new Set();
  // Match named exports of the form `export { a, b as c } from '...';`
  // and `export type { a, b } from '...';` and `export { a } from '...';`
  // (with multi-line forms).
  const pattern = /export\s+(?:type\s+)?\{([^}]+)\}\s*from\s*['"][^'"]+['"]/g;
  let m;
  while ((m = pattern.exec(text))) {
    for (const raw of m[1].split(',')) {
      const piece = raw.trim();
      if (!piece) continue;
      // Handle `a as b` (alias name is what's exported).
      const asIdx = piece.search(/\s+as\s+/);
      const exported = asIdx >= 0 ? piece.slice(asIdx).replace(/\s+as\s+/, '').trim() : piece;
      // Strip trailing comments / whitespace.
      const clean = exported.replace(/\/\/.*$/, '').trim();
      if (clean) symbols.add(clean);
    }
  }
  return symbols;
}

const symbolToSection = new Map();
for (const section of SECTIONS) {
  const indexPath = join(ROOT, 'src', section, 'index.ts');
  let text;
  try {
    text = readFileSync(indexPath, 'utf8');
  } catch {
    continue;
  }
  for (const sym of parseExports(text)) {
    if (!symbolToSection.has(sym)) symbolToSection.set(sym, section);
  }
}

// Walk tests/ and rewrite.
function walk(dir, out = []) {
  for (const name of readdirSync(dir)) {
    const full = join(dir, name);
    const st = statSync(full);
    if (st.isDirectory()) walk(full, out);
    else if (full.endsWith('.ts')) out.push(full);
  }
  return out;
}

const tests = walk(join(ROOT, 'tests'));

const importRe = /import\s+(?:type\s+)?\{([^}]+)\}\s+from\s+['"]((?:\.\.\/)+src\/index)['"];?/g;

let totalRewrites = 0;
const unmappedSymbols = new Set();

for (const file of tests) {
  const text = readFileSync(file, 'utf8');
  if (!text.includes('src/index')) continue;

  const newText = text.replace(importRe, (whole, names, src) => {
    const isType = /^\s*import\s+type\s/.test(whole);
    // Compute relative prefix (number of `../`).
    const upDirs = (src.match(/\.\.\//g) ?? []).length;
    const prefix = '../'.repeat(upDirs) + 'src/';

    // Build groups by section.
    const groups = new Map();
    for (const raw of names.split(',')) {
      const piece = raw.trim();
      if (!piece) continue;
      const cleaned = piece.replace(/\/\/.*$/, '').trim();
      if (!cleaned) continue;
      // Drop a leading `type ` modifier (per-symbol type-only imports).
      const isItemType = /^type\s+/.test(cleaned);
      const sym = cleaned.replace(/^type\s+/, '');
      const section = symbolToSection.get(sym);
      if (!section) {
        unmappedSymbols.add(sym);
        // Fall back to the worksheet/worksheet.ts in case (best-effort).
        continue;
      }
      const key = section;
      if (!groups.has(key)) groups.set(key, { type: [], value: [] });
      const bucket = groups.get(key);
      if (isItemType || isType) bucket.type.push(sym);
      else bucket.value.push(sym);
    }

    // Emit one import line per (section, kind).
    const lines = [];
    for (const section of [...groups.keys()].sort()) {
      const { type, value } = groups.get(section);
      const target = `'${prefix}${section}/index'`;
      if (value.length) lines.push(`import { ${value.sort().join(', ')} } from ${target};`);
      if (type.length) lines.push(`import type { ${type.sort().join(', ')} } from ${target};`);
    }
    return lines.join('\n');
  });

  if (newText !== text) {
    writeFileSync(file, newText);
    totalRewrites += 1;
    console.log(`rewrote: ${relative(ROOT, file)}`);
  }
}

console.log(`\nTotal files rewritten: ${totalRewrites}`);
if (unmappedSymbols.size) {
  console.log(`\nUnmapped symbols (left in place; investigate):`);
  for (const s of [...unmappedSymbols].sort()) console.log(`  - ${s}`);
}
