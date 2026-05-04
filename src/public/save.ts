// Public `saveWorkbook` entry point. Per docs/plan/05-read-write.md §1.2.
//
// **Stage 1 minimum**: emits the bare set of parts a Workbook needs to
// round-trip through `loadWorkbook`:
//
//     [Content_Types].xml
//     _rels/.rels
//     xl/workbook.xml
//     xl/_rels/workbook.xml.rels
//     xl/worksheets/sheetN.xml ...
//     xl/styles.xml
//     xl/sharedStrings.xml             (only when sst is non-empty)
//
// docProps / theme / VBA / drawings / charts are reserved for later
// iterations — load tolerates their absence.

import { chartToBytes } from '../chart/chart-xml';
import { chartExToBytes } from '../chart/cx/chartex-xml';
import type { Drawing, DrawingItem } from '../drawing/drawing';
import { drawingToBytes } from '../drawing/drawing-xml';
import type { XlsxSink } from '../io/sink';
import { corePropsToBytes } from '../packaging/core';
import { customPropsToBytes } from '../packaging/custom';
import { extendedPropsToBytes } from '../packaging/extended';
import { addDefault, addOverride, makeManifest, manifestToBytes } from '../packaging/manifest';
import { makeRelationships, type Relationships, relsToBytes } from '../packaging/relationships';
import { stylesheetToBytes } from '../styles/stylesheet-writer';
import { makeSharedStrings, sharedStringsToBytes } from '../workbook/shared-strings';
import type { Workbook } from '../workbook/workbook';
import type { LegacyComment } from '../worksheet/comments';
import { commentsToBytes, placeholderVmlDrawing } from '../worksheet/comments-xml';
import type { TableDefinition } from '../worksheet/table';
import { tableToBytes } from '../worksheet/table-xml';
import { worksheetToBytes } from '../worksheet/writer';
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
  ARC_WORKBOOK_RELS,
  CHARTEX_TYPE,
  CPROPS_TYPE,
  PACKAGE_WORKSHEETS,
  PKG_REL_NS,
  REL_NS,
  SHARED_STRINGS_TYPE,
  SHEET_MAIN_NS,
  STYLES_TYPE,
  THEME_TYPE,
  WORKSHEET_TYPE,
  XLSX_TYPE,
} from '../xml/namespaces';
import { createZipWriter } from '../zip/writer';

const CORE_PROPS_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml';
const EXT_PROPS_TYPE = 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
const CORE_PROPS_REL = `${PKG_REL_NS}/metadata/core-properties`;
const EXT_PROPS_REL = `${REL_NS}/extended-properties`;
const CUSTOM_PROPS_REL = `${REL_NS}/custom-properties`;
const THEME_REL = `${REL_NS}/theme`;
const TABLE_REL = `${REL_NS}/table`;
const TABLE_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const COMMENTS_REL = `${REL_NS}/comments`;
const VML_DRAWING_REL = `${REL_NS}/vmlDrawing`;
const COMMENTS_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
const VML_DRAWING_TYPE = 'application/vnd.openxmlformats-officedocument.vmlDrawing';
const DRAWING_REL = `${REL_NS}/drawing`;
const DRAWING_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const CHART_REL = `${REL_NS}/chart`;
const CHART_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

export interface SaveOptions {
  /** Reserved — passes through to the underlying ZIP writer when implemented. */
  compressionLevel?: number;
}

/** Convenience: serialise a Workbook to an in-memory `Uint8Array` xlsx. */
export async function workbookToBytes(wb: Workbook, opts: SaveOptions = {}): Promise<Uint8Array> {
  const { toBuffer } = await import('../io/node');
  const sink = toBuffer();
  await saveWorkbook(wb, sink, opts);
  return sink.result();
}

/** Save a workbook through the given sink. Returns once `finalize()` resolves. */
export async function saveWorkbook(wb: Workbook, sink: XlsxSink, _opts: SaveOptions = {}): Promise<void> {
  const writer = createZipWriter(sink);

  // ---- 1. assemble the per-sheet rels + serialise each worksheet ----------
  const sst = makeSharedStrings();
  interface SheetEmit {
    id: string;
    target: string;
    bytes: Uint8Array;
    /** Per-sheet rels — populated only if the worksheet has hyperlinks etc. */
    rels?: Relationships;
  }
  const sheetEmits: SheetEmit[] = [];
  // Workbook-global table counter so xl/tables/tableN.xml ids stay unique.
  const tableEmits: Array<{ id: number; bytes: Uint8Array }> = [];
  let nextTableId = 1;
  // Same counter pattern for comments parts. The comments part and the
  // VML drawing share their N because Excel always emits them paired.
  interface CommentEmit {
    id: number;
    commentsBytes: Uint8Array;
    vmlBytes: Uint8Array;
  }
  const commentEmits: CommentEmit[] = [];
  let nextCommentsId = 1;
  // Drawings: workbook-global drawingN counter for xl/drawings/drawingN.xml +
  // matching drawing-rels file (when chart items are present).
  interface DrawingEmit {
    id: number;
    bytes: Uint8Array;
    /** Drawing rels — when non-empty we also emit the rels file. */
    rels?: Relationships;
  }
  const drawingEmits: DrawingEmit[] = [];
  let nextDrawingId = 1;
  // Chart parts share a workbook-global counter so xl/charts/chartN.xml
  // ids stay unique.
  const chartEmits: Array<{ id: number; bytes: Uint8Array; isCx: boolean }> = [];
  let nextChartId = 1;
  wb.sheets.forEach((ref, i) => {
    const target = `worksheets/sheet${i + 1}.xml`;
    const sheetRels = makeRelationships();
    const registerTable = (table: TableDefinition): { rId: string } => {
      const tableId = nextTableId++;
      // Allocate a fresh worksheet-rels rId; numbering is local to this sheet.
      const rId = `rId${sheetRels.rels.length + 1}`;
      sheetRels.rels.push({
        id: rId,
        type: TABLE_REL,
        target: `../tables/table${tableId}.xml`,
      });
      // Emit the table part with its workbook-global id baked in.
      const xmlTable: TableDefinition = { ...table, id: tableId };
      tableEmits.push({ id: tableId, bytes: tableToBytes(xmlTable) });
      return { rId };
    };
    const registerComments = (comments: ReadonlyArray<LegacyComment>): { vmlRelId: string } => {
      const id = nextCommentsId++;
      const commentsRelId = `rId${sheetRels.rels.length + 1}`;
      sheetRels.rels.push({
        id: commentsRelId,
        type: COMMENTS_REL,
        target: `../comments${id}.xml`,
      });
      const vmlRelId = `rId${sheetRels.rels.length + 1}`;
      sheetRels.rels.push({
        id: vmlRelId,
        type: VML_DRAWING_REL,
        target: `../drawings/vmlDrawing${id}.vml`,
      });
      commentEmits.push({
        id,
        commentsBytes: commentsToBytes(comments),
        vmlBytes: placeholderVmlDrawing(),
      });
      return { vmlRelId };
    };
    const registerDrawing = (drawing: Drawing): { rId: string } => {
      const id = nextDrawingId++;
      const rId = `rId${sheetRels.rels.length + 1}`;
      sheetRels.rels.push({
        id: rId,
        type: DRAWING_REL,
        target: `../drawings/drawing${id}.xml`,
      });
      // Walk drawing items: for each chart with a payload, allocate a
      // workbook-global chartN id and a per-drawing rId. Emit the chart
      // part + collect a drawing-rels entry so drawing.xml's
      // <c:chart r:id> resolves.
      const drawingRels = makeRelationships();
      const itemsForXml: DrawingItem[] = [];
      for (const item of drawing.items) {
        if (item.content.kind === 'chart' && (item.content.chart.space || item.content.chart.cxSpace)) {
          const chartId = nextChartId++;
          const chartRId = `rId${drawingRels.rels.length + 1}`;
          drawingRels.rels.push({
            id: chartRId,
            type: CHART_REL,
            target: `../charts/chart${chartId}.xml`,
          });
          if (item.content.chart.cxSpace) {
            chartEmits.push({
              id: chartId,
              bytes: chartExToBytes(item.content.chart.cxSpace),
              isCx: true,
            });
          } else if (item.content.chart.space) {
            chartEmits.push({
              id: chartId,
              bytes: chartToBytes(item.content.chart.space),
              isCx: false,
            });
          }
          itemsForXml.push({
            anchor: item.anchor,
            content: { kind: 'chart', chart: { rId: chartRId } },
          });
        } else {
          itemsForXml.push(item);
        }
      }
      const emit: DrawingEmit = { id, bytes: drawingToBytes({ items: itemsForXml }) };
      if (drawingRels.rels.length > 0) emit.rels = drawingRels;
      drawingEmits.push(emit);
      return { rId };
    };
    const bytes = worksheetToBytes(ref.sheet, {
      sharedStrings: sst,
      rels: sheetRels,
      registerTable,
      registerComments,
      registerDrawing,
    });
    const emit: SheetEmit = { id: `rId${i + 1}`, target, bytes };
    if (sheetRels.rels.length > 0) emit.rels = sheetRels;
    sheetEmits.push(emit);
  });

  // ---- 2. workbook rels -- sheets first, then sst (if any), then styles ---
  // We pre-assigned rIds (rId1..rIdN) to the sheets so workbook.xml's <sheet
  // r:id> matches. Build the rels list directly rather than going through
  // appendRel (which auto-allocates ids).
  const wbRels = makeRelationships();
  wbRels.rels = sheetEmits.map((e) => ({
    id: e.id,
    type: `${REL_NS}/worksheet`,
    target: e.target,
  }));
  if (sst.entries.length > 0) {
    wbRels.rels.push({
      id: `rId${wbRels.rels.length + 1}`,
      type: `${REL_NS}/sharedStrings`,
      target: 'sharedStrings.xml',
    });
  }
  wbRels.rels.push({
    id: `rId${wbRels.rels.length + 1}`,
    type: `${REL_NS}/styles`,
    target: 'styles.xml',
  });
  if (wb.themeXml) {
    wbRels.rels.push({
      id: `rId${wbRels.rels.length + 1}`,
      type: THEME_REL,
      target: 'theme/theme1.xml',
    });
  }

  // ---- 3. workbook.xml ----------------------------------------------------
  const workbookXml = serializeWorkbookXml(
    wb,
    sheetEmits.map((e) => e.id),
  );
  await writer.addEntry(ARC_WORKBOOK, new TextEncoder().encode(workbookXml));
  await writer.addEntry(ARC_WORKBOOK_RELS, relsToBytes(wbRels));

  // ---- 4. each worksheet (and its rels file when present) ----------------
  for (let i = 0; i < sheetEmits.length; i++) {
    const e = sheetEmits[i];
    if (!e) continue;
    await writer.addEntry(`${PACKAGE_WORKSHEETS}/sheet${i + 1}.xml`, e.bytes);
    if (e.rels) {
      await writer.addEntry(`${PACKAGE_WORKSHEETS}/_rels/sheet${i + 1}.xml.rels`, relsToBytes(e.rels));
    }
  }

  // ---- 4b. table parts ----------------------------------------------------
  for (const t of tableEmits) {
    await writer.addEntry(`xl/tables/table${t.id}.xml`, t.bytes);
  }

  // ---- 4c. comments parts + matching VML drawings ------------------------
  for (const c of commentEmits) {
    await writer.addEntry(`xl/comments${c.id}.xml`, c.commentsBytes);
    await writer.addEntry(`xl/drawings/vmlDrawing${c.id}.vml`, c.vmlBytes);
  }

  // ---- 4d. drawings + their rels (when charts are embedded) -------------
  for (const d of drawingEmits) {
    await writer.addEntry(`xl/drawings/drawing${d.id}.xml`, d.bytes);
    if (d.rels) {
      await writer.addEntry(`xl/drawings/_rels/drawing${d.id}.xml.rels`, relsToBytes(d.rels));
    }
  }

  // ---- 4e. chart parts ---------------------------------------------------
  for (const c of chartEmits) {
    await writer.addEntry(`xl/charts/chart${c.id}.xml`, c.bytes);
  }

  // ---- 5. styles.xml + sharedStrings.xml (if any) -------------------------
  await writer.addEntry(ARC_STYLE, stylesheetToBytes(wb.styles));
  if (sst.entries.length > 0) {
    await writer.addEntry(ARC_SHARED_STRINGS, sharedStringsToBytes(sst));
  }

  // ---- 5b. theme1.xml (passthrough) — only when wb carries one. The theme
  // rel was already added to wbRels above so workbook.xml.rels references it.
  if (wb.themeXml) await writer.addEntry(ARC_THEME, wb.themeXml);

  // ---- 5c. docProps/{core,app,custom}.xml (when present) ------------------
  if (wb.properties) await writer.addEntry(ARC_CORE, corePropsToBytes(wb.properties));
  if (wb.appProperties) await writer.addEntry(ARC_APP, extendedPropsToBytes(wb.appProperties));
  if (wb.customProperties) await writer.addEntry(ARC_CUSTOM, customPropsToBytes(wb.customProperties));

  // ---- 6. root rels -------------------------------------------------------
  const rootRels: Relationships = { rels: [] };
  rootRels.rels.push({
    id: 'rId1',
    type: `${REL_NS}/officeDocument`,
    target: 'xl/workbook.xml',
  });
  if (wb.properties) {
    rootRels.rels.push({
      id: `rId${rootRels.rels.length + 1}`,
      type: CORE_PROPS_REL,
      target: 'docProps/core.xml',
    });
  }
  if (wb.appProperties) {
    rootRels.rels.push({
      id: `rId${rootRels.rels.length + 1}`,
      type: EXT_PROPS_REL,
      target: 'docProps/app.xml',
    });
  }
  if (wb.customProperties) {
    rootRels.rels.push({
      id: `rId${rootRels.rels.length + 1}`,
      type: CUSTOM_PROPS_REL,
      target: 'docProps/custom.xml',
    });
  }
  await writer.addEntry(ARC_ROOT_RELS, relsToBytes(rootRels));

  // ---- 7. [Content_Types].xml --------------------------------------------
  const manifest = makeManifest();
  addDefault(manifest, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');
  addDefault(manifest, 'xml', 'application/xml');
  addOverride(manifest, `/${ARC_WORKBOOK}`, XLSX_TYPE);
  for (let i = 0; i < sheetEmits.length; i++) {
    addOverride(manifest, `/${PACKAGE_WORKSHEETS}/sheet${i + 1}.xml`, WORKSHEET_TYPE);
  }
  addOverride(manifest, `/${ARC_STYLE}`, STYLES_TYPE);
  if (sst.entries.length > 0) {
    addOverride(manifest, `/${ARC_SHARED_STRINGS}`, SHARED_STRINGS_TYPE);
  }
  if (wb.themeXml) addOverride(manifest, `/${ARC_THEME}`, THEME_TYPE);
  for (const t of tableEmits) {
    addOverride(manifest, `/xl/tables/table${t.id}.xml`, TABLE_TYPE);
  }
  if (commentEmits.length > 0) {
    addDefault(manifest, 'vml', VML_DRAWING_TYPE);
  }
  for (const c of commentEmits) {
    addOverride(manifest, `/xl/comments${c.id}.xml`, COMMENTS_TYPE);
  }
  for (const d of drawingEmits) {
    addOverride(manifest, `/xl/drawings/drawing${d.id}.xml`, DRAWING_TYPE);
  }
  for (const c of chartEmits) {
    addOverride(manifest, `/xl/charts/chart${c.id}.xml`, c.isCx ? CHARTEX_TYPE : CHART_TYPE);
  }
  if (wb.properties) addOverride(manifest, `/${ARC_CORE}`, CORE_PROPS_TYPE);
  if (wb.appProperties) addOverride(manifest, `/${ARC_APP}`, EXT_PROPS_TYPE);
  if (wb.customProperties) addOverride(manifest, `/${ARC_CUSTOM}`, CPROPS_TYPE);
  await writer.addEntry(ARC_CONTENT_TYPES, manifestToBytes(manifest));

  // ---- 8. close ----------------------------------------------------------
  await writer.finalize();
}

/** Serialise the minimum `<workbook><sheets/></workbook>` Excel needs to load a sheet list. */
function serializeWorkbookXml(wb: Workbook, sheetRIds: ReadonlyArray<string>): string {
  const parts: string[] = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<workbook xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}">`,
    '<sheets>',
  ];
  wb.sheets.forEach((ref, i) => {
    const stateAttr = ref.state === 'visible' ? '' : ` state="${ref.state}"`;
    const rId = sheetRIds[i] ?? `rId${i + 1}`;
    parts.push(`<sheet name="${escapeAttr(ref.sheet.title)}" sheetId="${ref.sheetId}"${stateAttr} r:id="${rId}"/>`);
  });
  parts.push('</sheets>');
  if (wb.definedNames.length > 0) {
    parts.push('<definedNames>');
    for (const dn of wb.definedNames) {
      let attrs = ` name="${escapeAttr(dn.name)}"`;
      if (dn.scope !== undefined) attrs += ` localSheetId="${dn.scope}"`;
      if (dn.hidden) attrs += ' hidden="1"';
      if (dn.comment !== undefined) attrs += ` comment="${escapeAttr(dn.comment)}"`;
      parts.push(`<definedName${attrs}>${escapeText(dn.value)}</definedName>`);
    }
    parts.push('</definedNames>');
  }
  parts.push('</workbook>');
  return parts.join('');
}

const escapeText = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const escapeAttr = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
