// Shared-strings table read/write. Per docs/plan/05-read-write.md §4.
//
// Excel pulls every plain-string cell out of the sheet bodies and into
// `xl/sharedStrings.xml` so duplicates compress well. The format is a
// flat list of `<si>` entries, each holding either a single `<t>`
// (plain text) or a sequence of `<r><rPr/>?<t/></r>` runs (rich text).
//
// **Stage 1**: flat-string round-trip. Rich-text runs are concatenated
// into their plain-text body on read; rich-text fidelity round-trip is
// reserved for a later iteration of the loop. The wire format on write
// mirrors openpyxl's writer (single `<si><t>...</t></si>` per entry).
//
// The accumulator path is the part that runs during worksheet writes
// — `addSharedString(table, value)` is O(1) thanks to the in-place
// Map keyed by literal string.

import { escapeCellString, unescapeCellString } from '../utils/escape';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { SHEET_MAIN_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, findChildren, type XmlNode } from '../xml/tree';

const SST_TAG = `{${SHEET_MAIN_NS}}sst`;
const SI_TAG = `{${SHEET_MAIN_NS}}si`;
const T_TAG = `{${SHEET_MAIN_NS}}t`;
const R_TAG = `{${SHEET_MAIN_NS}}r`;

/**
 * Mutable shared-strings accumulator + lookup table. The same shape is
 * used during read (just populate `entries`) and write (call
 * `addSharedString` from the worksheet writer; emit to bytes at the end).
 */
export interface SharedStringsTable {
  /** Insertion-ordered list of unique entries. */
  entries: string[];
  /** Reverse lookup keyed by literal text. */
  index: Map<string, number>;
}

export function makeSharedStrings(): SharedStringsTable {
  return { entries: [], index: new Map() };
}

/**
 * Insert a string and return its index. Idempotent: calling with the
 * same value twice gives the same index. Empty strings are deduped just
 * like everything else.
 */
export function addSharedString(table: SharedStringsTable, value: string): number {
  const cached = table.index.get(value);
  if (cached !== undefined) return cached;
  const id = table.entries.length;
  table.entries.push(value);
  table.index.set(value, id);
  return id;
}

/** Look up a shared-string index by its literal text. Returns `undefined` for unknown values. */
export function getSharedStringIndex(table: SharedStringsTable, value: string): number | undefined {
  return table.index.get(value);
}

/** Read a shared-string by its 0-based index. Returns `undefined` for out-of-range. */
export function getSharedStringAt(table: SharedStringsTable, index: number): string | undefined {
  return table.entries[index];
}

/** Number of unique entries in the SST. */
export function sharedStringCount(table: SharedStringsTable): number {
  return table.entries.length;
}

// ---- read ------------------------------------------------------------------

/** Concatenate every `<t>` text node found inside an arbitrary XmlNode tree. */
const collectText = (node: XmlNode): string => {
  // Most common case: a direct `<si><t>x</t></si>` — bypass the recursion.
  if (node.children.length === 1) {
    const only = node.children[0];
    if (only && only.name === T_TAG) return unescapeCellString(only.text ?? '');
  }
  let out = '';
  for (const child of node.children) {
    if (child.name === T_TAG) {
      out += child.text ?? '';
    } else if (child.name === R_TAG) {
      const t = findChild(child, T_TAG);
      if (t?.text) out += t.text;
    }
  }
  return unescapeCellString(out);
};

/**
 * Parse a `xl/sharedStrings.xml` payload. Returns the table directly
 * (rather than just the array) so the worksheet writer can keep
 * appending to it without rebuilding the index.
 *
 * Rich-text runs are flattened to their plain-text body in this stage.
 */
export function parseSharedStringsXml(bytes: Uint8Array | string): SharedStringsTable {
  const root = parseXml(bytes);
  if (root.name !== SST_TAG) {
    throw new OpenXmlSchemaError(`parseSharedStringsXml: root is "${root.name}", expected sst`);
  }
  const table = makeSharedStrings();
  for (const si of findChildren(root, SI_TAG)) {
    const text = collectText(si);
    // Don't dedup — Excel preserves duplicate `<si>` entries by index, and
    // the worksheet `t="s"` references depend on slot, not on text equality.
    const id = table.entries.length;
    table.entries.push(text);
    if (!table.index.has(text)) table.index.set(text, id);
  }
  return table;
}

// ---- write -----------------------------------------------------------------

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Serialise a SharedStringsTable to its OOXML bytes. The `count`
 * attribute tracks total references (we don't know that here so we
 * report the same as `uniqueCount` — readers tolerate the discrepancy
 * and Excel ignores `count` in practice). `uniqueCount` always matches
 * `entries.length`.
 */
export function sharedStringsToBytes(table: SharedStringsTable): Uint8Array {
  return new TextEncoder().encode(serializeSharedStrings(table));
}

export function serializeSharedStrings(table: SharedStringsTable): string {
  const total = table.entries.length;
  const parts: string[] = [XML_HEADER, `<sst xmlns="${SHEET_MAIN_NS}" count="${total}" uniqueCount="${total}">`];
  for (const value of table.entries) {
    parts.push(serializeSi(value));
  }
  parts.push('</sst>');
  return parts.join('');
}

const escapeXmlText = (s: string): string =>
  // Reorder so '&' is replaced first; otherwise we'd double-escape ampersands
  // introduced by the later substitutions.
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const serializeSi = (value: string): string => {
  // Whitespace at either end needs xml:space="preserve" so Excel doesn't
  // collapse it. Mirrors openpyxl's emitter.
  const preserve = value.length > 0 && (value[0] === ' ' || value[value.length - 1] === ' ' || /[\t\n]/.test(value));
  const tAttr = preserve ? ' xml:space="preserve"' : '';
  return `<si><t${tAttr}>${escapeXmlText(escapeCellString(value))}</t></si>`;
};
