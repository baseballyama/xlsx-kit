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

import { findUserShapesRId, parseChartXml } from '../chart/chart-xml';
import { isChartExBytes, parseChartExXml } from '../chart/cx/chartex-xml';
import { parseUserShapesXml } from '../chart/user-shapes-xml';
import { parseChartsheetXml } from '../chartsheet/chartsheet-xml';
import { parseDrawingXml } from '../drawing/drawing-xml';
import { loadImage } from '../drawing/image';
import type { XlsxSource } from '../io/source';
import { corePropsFromBytes } from '../packaging/core';
import { customPropsFromBytes } from '../packaging/custom';
import { extendedPropsFromBytes } from '../packaging/extended';
import { manifestFromBytes } from '../packaging/manifest';
import { findById, relsFromBytes } from '../packaging/relationships';
import { parseStylesheetXml } from '../styles/stylesheet-reader';
import { OpenXmlSchemaError } from '../utils/exceptions';
import type { DefinedName } from '../workbook/defined-names';
import { makeDefinedName } from '../workbook/defined-names';
import { parseSharedStringsXml, type SharedStringsTable } from '../workbook/shared-strings';
import { createWorkbook, type SheetRef, type SheetState, type Workbook } from '../workbook/workbook';
import { parseCommentsXml } from '../worksheet/comments-xml';
import { parseWorksheetXml } from '../worksheet/reader';
import { parseTableXml } from '../worksheet/table-xml';
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
const DEFINED_NAMES_TAG = `{${SHEET_MAIN_NS}}definedNames`;
const DEFINED_NAME_TAG = `{${SHEET_MAIN_NS}}definedName`;
const RID_ATTR = `{${REL_NS}}id`;

/** Extract the `<definedNames>/<definedName>` entries from a parsed `xl/workbook.xml`. */
export function parseDefinedNames(workbookRoot: XmlNode): DefinedName[] {
  const wrapper = findChild(workbookRoot, DEFINED_NAMES_TAG);
  if (!wrapper) return [];
  const out: DefinedName[] = [];
  for (const node of findChildren(wrapper, DEFINED_NAME_TAG)) {
    const name = node.attrs['name'];
    if (!name) throw new OpenXmlSchemaError("workbook.xml: <definedName> is missing 'name'");
    const value = node.text ?? '';
    const opts: Partial<DefinedName> & { name: string; value: string } = { name, value };
    const scopeAttr = node.attrs['localSheetId'];
    if (scopeAttr !== undefined) {
      const scope = Number.parseInt(scopeAttr, 10);
      if (Number.isInteger(scope) && scope >= 0) opts.scope = scope;
    }
    if (node.attrs['hidden'] === '1' || node.attrs['hidden'] === 'true') opts.hidden = true;
    if (node.attrs['comment'] !== undefined) opts.comment = node.attrs['comment'];
    out.push(makeDefinedName(opts));
  }
  return out;
}

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
  const manifest = manifestFromBytes(archive.read(ARC_CONTENT_TYPES));

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
  const definedNamesFromXml = parseDefinedNames(wbRoot);

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
  if (definedNamesFromXml.length > 0) wb.definedNames = definedNamesFromXml;
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
    const loadTable = sheetRels
      ? (relId: string) => {
          const tRel = sheetRels.rels.find((r) => r.id === relId);
          if (!tRel) return undefined;
          const tablePath = resolveRelTarget(sheetPath, tRel.target);
          if (!archive.has(tablePath)) return undefined;
          return parseTableXml(archive.read(tablePath));
        }
      : undefined;
    const loadComments = sheetRels
      ? (relId: string) => {
          const cRel = sheetRels.rels.find((r) => r.id === relId);
          if (!cRel) return undefined;
          const cPath = resolveRelTarget(sheetPath, cRel.target);
          if (!archive.has(cPath)) return undefined;
          return parseCommentsXml(archive.read(cPath));
        }
      : undefined;
    const loadDrawing = sheetRels
      ? (relId: string) => {
          const dRel = sheetRels.rels.find((r) => r.id === relId);
          if (!dRel) return undefined;
          const dPath = resolveRelTarget(sheetPath, dRel.target);
          if (!archive.has(dPath)) return undefined;
          const drawing = parseDrawingXml(archive.read(dPath));
          // Phase-2: resolve drawing-rels to populate chart payloads.
          const dRelsPath = relsPathFor(dPath);
          if (archive.has(dRelsPath)) {
            const dRels = relsFromBytes(archive.read(dRelsPath));
            for (const item of drawing.items) {
              if (item.content.kind === 'chart') {
                const chartRId = item.content.chart.rId;
                if (!chartRId) continue;
                const chartRel = dRels.rels.find((r) => r.id === chartRId);
                if (!chartRel) continue;
                const chartPath = resolveRelTarget(dPath, chartRel.target);
                if (archive.has(chartPath)) {
                  const chartBytes = archive.read(chartPath);
                  if (isChartExBytes(chartBytes)) {
                    item.content.chart.cxSpace = parseChartExXml(chartBytes);
                  } else {
                    const space = parseChartXml(chartBytes);
                    // Resolve <c:userShapes r:id="..."> via the chart's
                    // own rels file (xl/charts/_rels/chartN.xml.rels).
                    const userShapesRId = findUserShapesRId(chartBytes);
                    if (userShapesRId) {
                      const chartRelsPath = relsPathFor(chartPath);
                      if (archive.has(chartRelsPath)) {
                        const chartRelsObj = relsFromBytes(archive.read(chartRelsPath));
                        const usRel = chartRelsObj.rels.find((r) => r.id === userShapesRId);
                        if (usRel) {
                          const usPath = resolveRelTarget(chartPath, usRel.target);
                          if (archive.has(usPath)) {
                            try {
                              space.userShapes = parseUserShapesXml(archive.read(usPath));
                            } catch {
                              // Tolerate parse failures (Excel sometimes
                              // emits chartDrawing parts with namespaces /
                              // shapes outside our model).
                            }
                          }
                        }
                      }
                    }
                    item.content.chart.space = space;
                  }
                }
              } else if (item.content.kind === 'picture') {
                const picRId = item.content.picture.rId;
                if (!picRId) continue;
                const picRel = dRels.rels.find((r) => r.id === picRId);
                if (!picRel) continue;
                const imgPath = resolveRelTarget(dPath, picRel.target);
                if (archive.has(imgPath)) {
                  try {
                    item.content.picture.image = loadImage(archive.read(imgPath));
                  } catch {
                    // Unknown / unsupported format — leave bytes-less; callers can read
                    // via the rId + archive directly if they need the raw payload.
                  }
                }
              }
            }
          }
          return drawing;
        }
      : undefined;
    // Distinguish worksheet vs chartsheet by inspecting the workbook-rels
    // entry's relationship type.
    const isChartsheet = rel.type === `${REL_NS}/chartsheet`;
    if (isChartsheet) {
      const chartsheet = parseChartsheetXml(archive.read(sheetPath), entry.name);
      // Inline drawing reference from the chartsheet XML.
      if (sheetRels) {
        // Find the drawing rel and resolve it the same way worksheets do.
        const drawingRel = sheetRels.rels.find((r) => r.type === `${REL_NS}/drawing`);
        if (drawingRel && loadDrawing) {
          const d = loadDrawing(drawingRel.id);
          if (d) chartsheet.drawing = d;
        }
      }
      const ref: SheetRef = {
        kind: 'chartsheet',
        sheet: chartsheet,
        sheetId: entry.sheetId,
        state: entry.state,
        rId: entry.rId,
      };
      wb.sheets.push(ref);
      continue;
    }
    const ws = parseWorksheetXml(archive.read(sheetPath), entry.name, {
      sharedStrings: sst,
      ...(sheetRels ? { rels: sheetRels } : {}),
      ...(loadTable ? { loadTable } : {}),
      ...(loadComments ? { loadComments } : {}),
      ...(loadDrawing ? { loadDrawing } : {}),
    });
    const ref: SheetRef = {
      kind: 'worksheet',
      sheet: ws,
      sheetId: entry.sheetId,
      state: entry.state,
      rId: entry.rId,
    };
    wb.sheets.push(ref);
  }

  captureWorkbookXmlExtras(wbRoot, wb);
  captureWorkbookRelsExtras(wbRels, wb);

  // Pass-through: capture parts we don't model (VBA / pivot / activeX /
  // OLE / customUI / customXml / etc.) so re-saving doesn't drop them.
  capturePassthrough(archive, manifest, wb);
  return wb;
}

/**
 * Walk top-level children of `<workbook>` and split anything that isn't
 * `<sheets>` or `<definedNames>` into the before/after halves the writer
 * inserts around the modeled elements. Order is preserved within each
 * half so things like `<fileVersion>`, `<workbookPr>`, `<bookViews>`,
 * `<calcPr>`, `<pivotCaches>`, `<extLst>` round-trip in document order.
 */
function captureWorkbookXmlExtras(wbRoot: XmlNode, wb: Workbook): void {
  const beforeSheets: XmlNode[] = [];
  const afterSheets: XmlNode[] = [];
  let seenSheets = false;
  for (const child of wbRoot.children) {
    if (child.name === SHEETS_TAG) {
      seenSheets = true;
      continue;
    }
    if (child.name === DEFINED_NAMES_TAG) continue;
    if (seenSheets) afterSheets.push(child);
    else beforeSheets.push(child);
  }
  if (beforeSheets.length > 0 || afterSheets.length > 0) {
    wb.workbookXmlExtras = { beforeSheets, afterSheets };
  }
}

/**
 * Capture workbook-rels entries that don't match a modeled type so the
 * writer can re-emit them with their original Id (and any captured
 * `<pivotCaches r:id="…"/>` etc. still resolves after a round-trip).
 * Modeled non-sheet rels (sst / styles / theme / vbaProject) keep their
 * original Id via `wb.workbookRelOriginalIds` so the writer can prefer
 * those over freshly allocated ones.
 */
function captureWorkbookRelsExtras(
  wbRels: import('../packaging/relationships').Relationships,
  wb: Workbook,
): void {
  const SHEET_RELS = new Set([`${REL_NS}/worksheet`, `${REL_NS}/chartsheet`]);
  const original: NonNullable<Workbook['workbookRelOriginalIds']> = {};
  const extras: Array<{ id: string; type: string; target: string }> = [];
  for (const rel of wbRels.rels) {
    if (SHEET_RELS.has(rel.type)) continue;
    if (rel.type === `${REL_NS}/sharedStrings`) {
      original.sharedStrings = rel.id;
      continue;
    }
    if (rel.type === `${REL_NS}/styles`) {
      original.styles = rel.id;
      continue;
    }
    if (rel.type === `${REL_NS}/theme`) {
      original.theme = rel.id;
      continue;
    }
    if (rel.type === `${REL_NS}/vbaProject`) {
      original.vbaProject = rel.id;
      continue;
    }
    extras.push({ id: rel.id, type: rel.type, target: rel.target });
  }
  if (Object.keys(original).length > 0) wb.workbookRelOriginalIds = original;
  if (extras.length > 0) wb.workbookRelsExtras = extras;
}

const PASSTHROUGH_PREFIXES: ReadonlyArray<string> = [
  'xl/activeX/',
  'xl/ctrlProps/',
  'xl/embeddings/',
  'xl/pivotCache/',
  'xl/pivotTables/',
  'xl/printerSettings/',
  'xl/queryTables/',
  'xl/slicerCaches/',
  'xl/slicers/',
  'customUI/',
  'customXml/',
];

const isControlVml = (path: string): boolean => {
  // Control VML drawings live next to chart drawings under xl/drawings/.
  // Comment VML (vmlDrawingN.vml) is regenerated from comments and must
  // NOT be captured as passthrough.
  if (!path.startsWith('xl/drawings/')) return false;
  if (!path.endsWith('.vml')) return false;
  return !path.includes('/vmlDrawing');
};

const isPassthroughPath = (path: string): boolean => {
  if (PASSTHROUGH_PREFIXES.some((p) => path.startsWith(p))) return true;
  return isControlVml(path);
};

/**
 * Walk the archive after the modeled parts are loaded and capture any
 * remaining content into `wb.passthrough`. The dedicated VBA project
 * binaries land on their own slots so the writer can promote the
 * workbook content type to xlsm.
 */
function capturePassthrough(
  archive: ZipArchive,
  manifest: import('../packaging/manifest').Manifest,
  wb: Workbook,
): void {
  const overrides = new Map<string, string>();
  for (const o of manifest.overrides) {
    // Manifest paths are package-absolute (`/xl/...`); strip the leading slash.
    overrides.set(o.partName.replace(/^\//, ''), o.contentType);
  }
  for (const path of archive.list()) {
    if (path === 'xl/vbaProject.bin') {
      wb.vbaProject = archive.read(path);
      continue;
    }
    if (path === 'xl/vbaProjectSignature.bin') {
      wb.vbaSignature = archive.read(path);
      continue;
    }
    if (!isPassthroughPath(path)) continue;
    if (!wb.passthrough) wb.passthrough = new Map();
    wb.passthrough.set(path, archive.read(path));
    const ct = overrides.get(path);
    if (ct !== undefined) {
      if (!wb.passthroughContentTypes) wb.passthroughContentTypes = new Map();
      wb.passthroughContentTypes.set(path, ct);
    }
  }
}
