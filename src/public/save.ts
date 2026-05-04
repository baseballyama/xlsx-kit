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
import { userShapesToBytes } from '../chart/user-shapes-xml';
import { chartsheetToBytes } from '../chartsheet/chartsheet-xml';
import type { Drawing, DrawingItem } from '../drawing/drawing';
import { drawingToBytes } from '../drawing/drawing-xml';
import { IMAGE_FORMAT_EXTENSION, IMAGE_FORMAT_MIME, type XlsxImageFormat } from '../drawing/image';
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
import { serializeXml as serializeXmlNode } from '../xml/serializer';
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
const IMAGE_REL = `${REL_NS}/image`;
const CHARTSHEET_REL = `${REL_NS}/chartsheet`;
const CHARTSHEET_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml';
const CHART_USER_SHAPES_REL = `${REL_NS}/chartUserShapes`;

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
    /** OOXML relationship type used in workbook.xml.rels. */
    relType: string;
    /** ZIP archive path the bytes are written to. */
    archivePath: string;
    /** Override content type for [Content_Types].xml. */
    contentType: string;
    /** Per-sheet rels — populated only if the sheet has hyperlinks / drawings / etc. */
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
  const chartEmits: Array<{
    id: number;
    bytes: Uint8Array;
    isCx: boolean;
    /** Per-chart rels file (only emitted when chart.userShapes is set). */
    rels?: Relationships;
  }> = [];
  let nextChartId = 1;
  // Workbook-global counter for chartDrawing parts. Each chart with
  // userShapes set produces one xl/drawings/chartDrawingN.xml entry.
  const userShapeEmits: Array<{ id: number; bytes: Uint8Array }> = [];
  let nextUserShapesId = 1;
  // Workbook-global counter for image media parts. Excel uses
  // xl/media/imageN.{ext} where N is shared across the package.
  interface ImageEmit {
    id: number;
    /** File extension (without the dot) — also drives the manifest Default. */
    ext: string;
    bytes: Uint8Array;
  }
  const imageEmits: ImageEmit[] = [];
  const imageExts = new Set<string>();
  let nextImageId = 1;
  // Track separate counters so worksheet / chartsheet IDs do not collide
  // (Excel uses xl/worksheets/sheetN.xml and xl/chartsheets/sheetM.xml
  // independently of one another).
  let nextWorksheetId = 1;
  let nextChartsheetId = 1;
  // Pre-claim every original rId so freshly allocated ones (sheets without
  // a captured rId, plus modeled non-sheet rels) avoid collision with
  // captured workbookRelsExtras and the modeled non-sheet original ids.
  const claimedRIds = new Set<string>();
  for (const ref of wb.sheets) {
    if (ref.rId !== undefined) claimedRIds.add(ref.rId);
  }
  if (wb.workbookRelOriginalIds) {
    for (const v of Object.values(wb.workbookRelOriginalIds)) {
      if (typeof v === 'string') claimedRIds.add(v);
    }
  }
  if (wb.workbookRelsExtras) {
    for (const e of wb.workbookRelsExtras) claimedRIds.add(e.id);
  }
  let nextRIdCursor = 1;
  const allocateRId = (): string => {
    let id = `rId${nextRIdCursor}`;
    while (claimedRIds.has(id)) {
      nextRIdCursor++;
      id = `rId${nextRIdCursor}`;
    }
    claimedRIds.add(id);
    nextRIdCursor++;
    return id;
  };
  wb.sheets.forEach((ref, _i) => {
    const isChartsheet = ref.kind === 'chartsheet';
    const target = isChartsheet ? `chartsheets/sheet${nextChartsheetId}.xml` : `worksheets/sheet${nextWorksheetId}.xml`;
    const archivePath = isChartsheet
      ? `xl/chartsheets/sheet${nextChartsheetId}.xml`
      : `xl/worksheets/sheet${nextWorksheetId}.xml`;
    if (isChartsheet) nextChartsheetId++;
    else nextWorksheetId++;
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
            const space = item.content.chart.space;
            // If the chart carries user shapes, allocate the chartDrawingN
            // part + a per-chart rels file referencing it, then bake the
            // resulting r:id into the chart's <c:userShapes> element.
            let userShapesRId: string | undefined;
            let chartRelsFile: Relationships | undefined;
            if (space.userShapes && space.userShapes.shapes.length > 0) {
              const userShapesId = nextUserShapesId++;
              userShapesRId = 'rId1';
              chartRelsFile = makeRelationships();
              chartRelsFile.rels.push({
                id: userShapesRId,
                type: CHART_USER_SHAPES_REL,
                target: `../drawings/chartDrawing${userShapesId}.xml`,
              });
              userShapeEmits.push({
                id: userShapesId,
                bytes: userShapesToBytes(space.userShapes),
              });
            }
            const chartEmit: typeof chartEmits[number] = {
              id: chartId,
              bytes: chartToBytes(space, userShapesRId !== undefined ? { userShapesRId } : {}),
              isCx: false,
            };
            if (chartRelsFile) chartEmit.rels = chartRelsFile;
            chartEmits.push(chartEmit);
          }
          itemsForXml.push({
            anchor: item.anchor,
            content: { kind: 'chart', chart: { rId: chartRId } },
          });
        } else if (item.content.kind === 'picture' && item.content.picture.image) {
          const img = item.content.picture.image;
          const ext = IMAGE_FORMAT_EXTENSION[img.format];
          const imageId = nextImageId++;
          const picRId = `rId${drawingRels.rels.length + 1}`;
          drawingRels.rels.push({
            id: picRId,
            type: IMAGE_REL,
            target: `../media/image${imageId}.${ext}`,
          });
          imageEmits.push({ id: imageId, ext, bytes: img.bytes });
          imageExts.add(ext);
          itemsForXml.push({
            anchor: item.anchor,
            content: {
              kind: 'picture',
              picture: {
                rId: picRId,
                ...(item.content.picture.name !== undefined ? { name: item.content.picture.name } : {}),
                ...(item.content.picture.descr !== undefined ? { descr: item.content.picture.descr } : {}),
                ...(item.content.picture.hidden !== undefined ? { hidden: item.content.picture.hidden } : {}),
                ...(item.content.picture.spPr ? { spPr: item.content.picture.spPr } : {}),
              },
            },
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
    let bytes: Uint8Array;
    if (ref.kind === 'worksheet') {
      bytes = worksheetToBytes(ref.sheet, {
        sharedStrings: sst,
        rels: sheetRels,
        registerTable,
        registerComments,
        registerDrawing,
      });
    } else {
      // Chartsheet: register the drawing (if any), then emit the
      // chartsheet part with the resulting r:id baked in.
      let drawingRId: string | undefined;
      if (ref.sheet.drawing) drawingRId = registerDrawing(ref.sheet.drawing).rId;
      bytes = chartsheetToBytes(ref.sheet, drawingRId !== undefined ? { drawingRId } : {});
    }
    const sheetRId = ref.rId ?? allocateRId();
    const emit: SheetEmit = {
      id: sheetRId,
      target,
      bytes,
      relType: isChartsheet ? CHARTSHEET_REL : `${REL_NS}/worksheet`,
      archivePath,
      contentType: isChartsheet ? CHARTSHEET_TYPE : WORKSHEET_TYPE,
    };
    if (sheetRels.rels.length > 0) emit.rels = sheetRels;
    sheetEmits.push(emit);
  });

  // ---- 2. workbook rels -- sheets first, then sst (if any), then styles,
  // then theme / vbaProject, then any captured workbookRelsExtras (e.g.
  // pivotCacheDefinition rels referenced by `<pivotCaches>`).
  // Modeled non-sheet rels prefer the rId captured at load time so any
  // captured extras XML using that Id still resolves after the round-trip.
  const wbRels = makeRelationships();
  wbRels.rels = sheetEmits.map((e) => ({
    id: e.id,
    type: e.relType,
    target: e.target,
  }));
  const orig = wb.workbookRelOriginalIds;
  if (sst.entries.length > 0) {
    wbRels.rels.push({
      id: orig?.sharedStrings ?? allocateRId(),
      type: `${REL_NS}/sharedStrings`,
      target: 'sharedStrings.xml',
    });
  }
  wbRels.rels.push({
    id: orig?.styles ?? allocateRId(),
    type: `${REL_NS}/styles`,
    target: 'styles.xml',
  });
  if (wb.themeXml) {
    wbRels.rels.push({
      id: orig?.theme ?? allocateRId(),
      type: THEME_REL,
      target: 'theme/theme1.xml',
    });
  }
  if (wb.vbaProject) {
    wbRels.rels.push({
      id: orig?.vbaProject ?? allocateRId(),
      type: `${REL_NS}/vbaProject`,
      target: 'vbaProject.bin',
    });
  }
  if (wb.workbookRelsExtras) {
    for (const e of wb.workbookRelsExtras) {
      wbRels.rels.push({ id: e.id, type: e.type, target: e.target });
    }
  }

  // ---- 3. workbook.xml ----------------------------------------------------
  const workbookXml = serializeWorkbookXml(
    wb,
    sheetEmits.map((e) => e.id),
  );
  await writer.addEntry(ARC_WORKBOOK, new TextEncoder().encode(workbookXml));
  await writer.addEntry(ARC_WORKBOOK_RELS, relsToBytes(wbRels));

  // ---- 4. each worksheet / chartsheet (and its rels file when present) --
  for (const e of sheetEmits) {
    await writer.addEntry(e.archivePath, e.bytes);
    if (e.rels) {
      // The rels file sits alongside its part: `xl/<dir>/_rels/<file>.rels`.
      const slash = e.archivePath.lastIndexOf('/');
      const dir = e.archivePath.slice(0, slash);
      const file = e.archivePath.slice(slash + 1);
      await writer.addEntry(`${dir}/_rels/${file}.rels`, relsToBytes(e.rels));
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

  // ---- 4e. chart parts (and per-chart rels when user shapes attach) -----
  for (const c of chartEmits) {
    await writer.addEntry(`xl/charts/chart${c.id}.xml`, c.bytes);
    if (c.rels) {
      await writer.addEntry(`xl/charts/_rels/chart${c.id}.xml.rels`, relsToBytes(c.rels));
    }
  }

  // ---- 4f. user-shape drawings (xl/drawings/chartDrawingN.xml) ----------
  for (const us of userShapeEmits) {
    await writer.addEntry(`xl/drawings/chartDrawing${us.id}.xml`, us.bytes);
  }

  // ---- 4g. embedded images (xl/media/imageN.{ext}) ----------------------
  for (const img of imageEmits) {
    await writer.addEntry(`xl/media/image${img.id}.${img.ext}`, img.bytes);
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

  // ---- 5d. VBA project + signature (when present) -------------------------
  if (wb.vbaProject) await writer.addEntry('xl/vbaProject.bin', wb.vbaProject);
  if (wb.vbaSignature) await writer.addEntry('xl/vbaProjectSignature.bin', wb.vbaSignature);

  // ---- 5e. Pass-through bytes (pivot / activeX / customXml / etc.) -------
  if (wb.passthrough) {
    for (const [path, bytes] of wb.passthrough) await writer.addEntry(path, bytes);
  }

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
  // VBA-bearing workbooks promote the workbook content type to xlsm.
  const workbookContentType = wb.vbaProject
    ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
    : XLSX_TYPE;
  addOverride(manifest, `/${ARC_WORKBOOK}`, workbookContentType);
  for (const e of sheetEmits) {
    addOverride(manifest, `/${e.archivePath}`, e.contentType);
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
  for (const us of userShapeEmits) {
    addOverride(manifest, `/xl/drawings/chartDrawing${us.id}.xml`, DRAWING_TYPE);
  }
  if (wb.vbaProject) {
    addDefault(manifest, 'bin', 'application/vnd.ms-office.vbaProject');
  }
  if (wb.passthrough) {
    for (const path of wb.passthrough.keys()) {
      const ct = wb.passthroughContentTypes?.get(path);
      if (ct !== undefined) addOverride(manifest, `/${path}`, ct);
    }
  }
  // Each unique image extension gets a Default entry (`<Default Extension="png" ContentType="image/png"/>`).
  for (const ext of imageExts) {
    const fmt = (Object.entries(IMAGE_FORMAT_EXTENSION).find(([, e]) => e === ext) ?? [])[0] as
      | XlsxImageFormat
      | undefined;
    if (fmt) addDefault(manifest, ext, IMAGE_FORMAT_MIME[fmt]);
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
  ];
  if (wb.workbookXmlExtras?.beforeSheets) {
    for (const node of wb.workbookXmlExtras.beforeSheets) parts.push(serializeChildNode(node));
  }
  parts.push('<sheets>');
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
  if (wb.workbookXmlExtras?.afterSheets) {
    for (const node of wb.workbookXmlExtras.afterSheets) parts.push(serializeChildNode(node));
  }
  parts.push('</workbook>');
  return parts.join('');
}

const escapeText = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const escapeAttr = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

/**
 * Serialise an XmlNode child of `<workbook>` for inline injection back
 * into the workbook XML stream. Reuses serializeXml then strips the
 * declaration. Captured nodes carry Clark-notation names so namespace
 * prefixes get reallocated by serializeXml — Excel tolerates the extra
 * `xmlns="…"` declarations on each captured root.
 */
function serializeChildNode(node: import('../xml/tree').XmlNode): string {
  const bytes = serializeXmlNode(node, { xmlDeclaration: false });
  return new TextDecoder().decode(bytes);
}
