// Public `loadWorkbook` entry point. Per docs/plan/05-read-write.md §1.1.
//
// **This is the minimum-skeleton stage**: open zip → parse manifest →
// resolve workbook part path → parse the `<sheets>` list → for each
// sheet, allocate an empty Worksheet (title + sheetId + state). Reading
// the actual cell content / styles / sharedStrings / theme / docProps
// happens in the next iterations of the loop.
//
// The skeleton is enough to round-trip through openpyxl's
// `genuine/empty.xlsx` fixture (3 empty sheets) and to give the rest of
// phase 3 a stable scaffolding to layer onto.

import type { XlsxSource } from '../io/source';
import { corePropsFromBytes } from '../packaging/core';
import { customPropsFromBytes } from '../packaging/custom';
import { extendedPropsFromBytes } from '../packaging/extended';
import { manifestFromBytes } from '../packaging/manifest';
import { findById, relsFromBytes } from '../packaging/relationships';
import { parseStylesheetXml } from '../styles/stylesheet-reader';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { parseSharedStringsXml, type SharedStringsTable } from '../workbook/shared-strings';
import { createWorkbook, type SheetRef, type SheetState, type Workbook } from '../workbook/workbook';
import { parseWorksheetXml } from '../worksheet/reader';
import {
  ARC_APP,
  ARC_CONTENT_TYPES,
  ARC_CORE,
  ARC_CUSTOM,
  ARC_ROOT_RELS,
  ARC_SHARED_STRINGS,
  ARC_STYLE,
  ARC_THEME,
  ARC_WORKBOOK,
  parseQName,
  REL_NS,
  SHEET_MAIN_NS,
} from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, findChildren, type XmlNode } from '../xml/tree';
import { openZip, type ZipArchive } from '../zip/reader';

/** Options for {@link loadWorkbook}. The full surface lands in later iterations. */
export interface LoadOptions {
  /** Reserved — not yet implemented. */
  readOnly?: boolean;
  /** Reserved — not yet implemented. */
  keepLinks?: boolean;
  /** Reserved — not yet implemented. */
  keepVba?: boolean;
  /** Reserved — not yet implemented. */
  dataOnly?: boolean;
  /** Reserved — not yet implemented. */
  richText?: boolean;
}

/** Office Document relationship type — the package-root pointer to `xl/workbook.xml`. */
const OFFICE_DOC_REL_TYPE = `${REL_NS}/officeDocument`;

/**
 * Resolve an OPC relationship target against its source part path.
 *
 * - Targets starting with `/` are package-absolute.
 * - Otherwise the target is relative to the source part's parent directory.
 * - `..` segments collapse normally.
 */
export function resolveRelTarget(sourcePartPath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const lastSlash = sourcePartPath.lastIndexOf('/');
  const parentDir = lastSlash >= 0 ? sourcePartPath.slice(0, lastSlash + 1) : '';
  const joined = parentDir + target;
  return normalizePath(joined);
}

function normalizePath(path: string): string {
  const segments = path.split('/');
  const out: string[] = [];
  for (const seg of segments) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') {
      out.pop();
      continue;
    }
    out.push(seg);
  }
  return out.join('/');
}

/** Sibling rels-part path for a given part. `xl/workbook.xml` → `xl/_rels/workbook.xml.rels`. */
function relsPathFor(partPath: string): string {
  const i = partPath.lastIndexOf('/');
  if (i < 0) return `_rels/${partPath}.rels`;
  return `${partPath.slice(0, i)}/_rels/${partPath.slice(i + 1)}.rels`;
}

interface SheetEntry {
  /** Display name (`sheet/@name`). */
  name: string;
  /** Workbook-scope sheetId (`sheet/@sheetId`). */
  sheetId: number;
  /** Workbook rels Id (`sheet/@r:id`). */
  rId: string;
  /** Visibility state — defaults to `'visible'` when the attribute is absent. */
  state: SheetState;
}

const SHEET_TAG = `{${SHEET_MAIN_NS}}sheet`;
const SHEETS_TAG = `{${SHEET_MAIN_NS}}sheets`;
const RID_ATTR = `{${REL_NS}}id`;

/** Extract the `<sheets>/<sheet>` entries from a parsed `xl/workbook.xml`. */
export function parseSheetEntries(workbookRoot: XmlNode): SheetEntry[] {
  const sheets = findChild(workbookRoot, SHEETS_TAG);
  if (!sheets) return [];
  const out: SheetEntry[] = [];
  for (const node of findChildren(sheets, SHEET_TAG)) {
    const name = node.attrs['name'];
    const sheetIdAttr = node.attrs['sheetId'];
    const rId = node.attrs[RID_ATTR];
    if (!name) throw new OpenXmlSchemaError("workbook.xml: <sheet> is missing 'name'");
    if (!sheetIdAttr) throw new OpenXmlSchemaError(`workbook.xml: <sheet name="${name}"> is missing 'sheetId'`);
    if (!rId) throw new OpenXmlSchemaError(`workbook.xml: <sheet name="${name}"> is missing 'r:id'`);
    const sheetId = Number.parseInt(sheetIdAttr, 10);
    if (!Number.isInteger(sheetId) || sheetId < 1) {
      throw new OpenXmlSchemaError(
        `workbook.xml: <sheet name="${name}"> sheetId "${sheetIdAttr}" is not a positive integer`,
      );
    }
    const stateAttr = node.attrs['state'];
    let state: SheetState = 'visible';
    if (stateAttr === 'hidden' || stateAttr === 'veryHidden') state = stateAttr;
    out.push({ name, sheetId, rId, state });
  }
  return out;
}

/**
 * Load a workbook from any {@link XlsxSource}. Currently produces a
 * scaffold Workbook: each Worksheet is empty (no cells / styles /
 * shared strings / theme yet). The next phase-3 iterations layer those
 * in atop the same skeleton.
 */
export async function loadWorkbook(source: XlsxSource, _opts: LoadOptions = {}): Promise<Workbook> {
  const archive = await openZip(source);
  try {
    return loadWorkbookFromArchive(archive);
  } finally {
    archive.close();
  }
}

/** Internal: same as {@link loadWorkbook} but operating on an already-opened archive. */
function loadWorkbookFromArchive(archive: ZipArchive): Workbook {
  // 1. Manifest — resolves which override entries the package declares.
  if (!archive.has(ARC_CONTENT_TYPES)) {
    throw new OpenXmlSchemaError(`loadWorkbook: missing "${ARC_CONTENT_TYPES}"`);
  }
  manifestFromBytes(archive.read(ARC_CONTENT_TYPES));

  // 2. Root rels → resolve the office-document relationship to the workbook part path.
  if (!archive.has(ARC_ROOT_RELS)) {
    throw new OpenXmlSchemaError(`loadWorkbook: missing "${ARC_ROOT_RELS}"`);
  }
  const rootRels = relsFromBytes(archive.read(ARC_ROOT_RELS));
  const officeRel = rootRels.rels.find((r) => r.type === OFFICE_DOC_REL_TYPE);
  if (!officeRel) {
    throw new OpenXmlSchemaError('loadWorkbook: root rels missing officeDocument relationship');
  }
  const workbookPath = resolveRelTarget('', officeRel.target);
  if (workbookPath !== ARC_WORKBOOK) {
    // Most xlsx files put the workbook at xl/workbook.xml. We accept any
    // path the rels point at as long as the archive holds it.
    if (!archive.has(workbookPath)) {
      throw new OpenXmlSchemaError(`loadWorkbook: workbook part "${workbookPath}" not found in archive`);
    }
  }

  // 3. workbook.xml — parse to extract sheet metadata only.
  const wbRoot = parseXml(archive.read(workbookPath));
  if (parseQName(wbRoot.name).local !== 'workbook') {
    throw new OpenXmlSchemaError(`loadWorkbook: ${workbookPath} root is "${wbRoot.name}", expected workbook`);
  }
  const sheetEntries = parseSheetEntries(wbRoot);

  // 4. workbook.xml.rels — needed to resolve each sheet's rId to a part path.
  const wbRelsPath = relsPathFor(workbookPath);
  // openpyxl tolerates a missing workbook.xml.rels (it implies no sheets); in
  // practice every Excel file has one. We require it so a malformed package
  // surfaces as an OpenXmlSchemaError, not a silently-empty workbook.
  if (sheetEntries.length > 0 && !archive.has(wbRelsPath)) {
    throw new OpenXmlSchemaError(`loadWorkbook: workbook has sheets but rels part "${wbRelsPath}" is missing`);
  }
  const wbRels = archive.has(wbRelsPath) ? relsFromBytes(archive.read(wbRelsPath)) : { rels: [] };

  // 4b. sharedStrings.xml — optional. The workbook rels can also point at a
  // non-default location; for the minimum-skeleton stage we look at the
  // canonical `xl/sharedStrings.xml` path first, then fall back to the rels
  // entry if present.
  let sharedStrings: SharedStringsTable | undefined;
  if (archive.has(ARC_SHARED_STRINGS)) {
    sharedStrings = parseSharedStringsXml(archive.read(ARC_SHARED_STRINGS));
  } else {
    const sstRel = wbRels.rels.find((r) => r.type === `${REL_NS}/sharedStrings`);
    if (sstRel) {
      const sstPath = resolveRelTarget(workbookPath, sstRel.target);
      if (archive.has(sstPath)) {
        sharedStrings = parseSharedStringsXml(archive.read(sstPath));
      }
    }
  }
  const sst: ReadonlyArray<string> = sharedStrings?.entries ?? [];

  // 4c. styles.xml — optional. Same default-or-rels lookup as sst.
  let styles: ReturnType<typeof parseStylesheetXml> | undefined;
  if (archive.has(ARC_STYLE)) {
    styles = parseStylesheetXml(archive.read(ARC_STYLE));
  } else {
    const stylesRel = wbRels.rels.find((r) => r.type === `${REL_NS}/styles`);
    if (stylesRel) {
      const stylesPath = resolveRelTarget(workbookPath, stylesRel.target);
      if (archive.has(stylesPath)) {
        styles = parseStylesheetXml(archive.read(stylesPath));
      }
    }
  }

  // 4d. docProps/{core,app,custom}.xml — package-level metadata. Each part is
  // optional; absent ones leave the matching Workbook field undefined. We
  // walk both the canonical path and the root rels so non-default layouts
  // (rare but legal) still resolve.
  const properties = archive.has(ARC_CORE) ? corePropsFromBytes(archive.read(ARC_CORE)) : undefined;
  const appProperties = archive.has(ARC_APP) ? extendedPropsFromBytes(archive.read(ARC_APP)) : undefined;
  const customProperties = archive.has(ARC_CUSTOM) ? customPropsFromBytes(archive.read(ARC_CUSTOM)) : undefined;

  // 4e. xl/theme/theme1.xml — kept verbatim. Excel renders with this exact
  // payload; round-tripping the bytes avoids drift.
  const themeXml: Uint8Array | undefined = (() => {
    if (archive.has(ARC_THEME)) return archive.read(ARC_THEME);
    const themeRel = wbRels.rels.find((r) => r.type === `${REL_NS}/theme`);
    if (themeRel) {
      const themePath = resolveRelTarget(workbookPath, themeRel.target);
      if (archive.has(themePath)) return archive.read(themePath);
    }
    return undefined;
  })();

  // 5. Build the Workbook. We bypass `addWorksheet` because that allocates
  // sheetIds via `allocateSheetId`; load preserves the IDs from XML.
  const wb = createWorkbook();
  if (styles) wb.styles = styles;
  if (properties) wb.properties = properties;
  if (appProperties) wb.appProperties = appProperties;
  if (customProperties) wb.customProperties = customProperties;
  if (themeXml) wb.themeXml = themeXml;
  const seenTitles = new Set<string>();
  for (const entry of sheetEntries) {
    if (seenTitles.has(entry.name)) {
      throw new OpenXmlSchemaError(`loadWorkbook: duplicate sheet name "${entry.name}"`);
    }
    seenTitles.add(entry.name);
    const rel = findById(wbRels, entry.rId);
    if (!rel) {
      throw new OpenXmlSchemaError(`loadWorkbook: sheet "${entry.name}" rId "${entry.rId}" has no matching rels entry`);
    }
    const sheetPath = resolveRelTarget(workbookPath, rel.target);
    if (!archive.has(sheetPath)) {
      throw new OpenXmlSchemaError(`loadWorkbook: sheet part "${sheetPath}" not found in archive`);
    }
    const sheetRelsPath = relsPathFor(sheetPath);
    const sheetRels = archive.has(sheetRelsPath) ? relsFromBytes(archive.read(sheetRelsPath)) : undefined;
    const ws = parseWorksheetXml(archive.read(sheetPath), entry.name, {
      sharedStrings: sst,
      ...(sheetRels ? { rels: sheetRels } : {}),
    });
    const ref: SheetRef = { kind: 'worksheet', sheet: ws, sheetId: entry.sheetId, state: entry.state };
    wb.sheets.push(ref);
  }
  return wb;
}
