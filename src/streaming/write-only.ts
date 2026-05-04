// Streaming write-only workbook. Per docs/plan/06-streaming.md §3.
//
// `createWriteOnlyWorkbook` lets callers append rows one at a time
// without holding a full Workbook in memory. The current backend uses
// the buffered ZIP writer underneath — so the API is correct but the
// 100M-cells / 1GB-heap acceptance criterion isn't met yet (that needs
// a streaming deflate writer; see TODO at the bottom). The API surface
// matches the doc spec so the perf rewrite is a drop-in later.

import type { CellValue } from '../cell/cell';
import type { XlsxSink } from '../io/sink';
import { addOverride, makeManifest, manifestToBytes } from '../packaging/manifest';
import { makeRelationships, relsToBytes } from '../packaging/relationships';
import {
  addBorder,
  addCellXf,
  addFill,
  addFont,
  addNumFmt,
  type CellXf,
  defaultCellXf,
  makeStylesheet,
  type Stylesheet,
} from '../styles/stylesheet';
import { stylesheetToBytes } from '../styles/stylesheet-writer';
import type { Alignment } from '../styles/alignment';
import type { Border } from '../styles/borders';
import type { Fill } from '../styles/fills';
import type { Font } from '../styles/fonts';
import type { Protection } from '../styles/protection';
import { OpenXmlIoError } from '../utils/exceptions';
import { makeSharedStrings, sharedStringsToBytes } from '../workbook/shared-strings';
import { worksheetToBytes } from '../worksheet/writer';
import { makeWorksheet, setCell, type Worksheet } from '../worksheet/worksheet';
import {
  ARC_CONTENT_TYPES,
  ARC_ROOT_RELS,
  ARC_SHARED_STRINGS,
  ARC_STYLE,
  ARC_WORKBOOK,
  ARC_WORKBOOK_RELS,
  PKG_REL_NS,
  REL_NS,
  SHARED_STRINGS_TYPE,
  SHEET_MAIN_NS,
  STYLES_TYPE,
  WORKSHEET_TYPE,
  XLSX_TYPE,
} from '../xml/namespaces';
import { createZipWriter } from '../zip/writer';

const escapeAttr = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export interface WriteOnlyOptions {
  /** Reserved — currently ignored (the buffered backend doesn't honour it). */
  estimatedMaxRow?: number;
}

export interface WriteOnlyStyle {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  numberFormat?: string;
  protection?: Protection;
}

export type WriteOnlyRowItem = CellValue | { value: CellValue; style?: WriteOnlyStyle };

export interface WriteOnlyWorksheet {
  readonly title: string;
  appendRow(row: WriteOnlyRowItem[]): Promise<void>;
  setColumnWidth(col: number, width: number): void;
  close(): Promise<void>;
}

export interface WriteOnlyWorkbook {
  /** Add a new worksheet. The previous worksheet must be `close()`d first. */
  addWorksheet(title: string): Promise<WriteOnlyWorksheet>;
  /** Finalise the archive: emits styles / sharedStrings / workbook / manifest / rels. */
  finalize(): Promise<void>;
}

const validateTitle = (title: string, taken: Set<string>): void => {
  if (typeof title !== 'string' || title.length === 0 || title.length > 31) {
    throw new OpenXmlIoError(`Worksheet title must be 1..31 chars; got "${title}"`);
  }
  if (taken.has(title)) {
    throw new OpenXmlIoError(`Worksheet title "${title}" is already in use`);
  }
};

/** Allocate a CellXf id for a style spec. Mirrors cell-style.ts but works directly on the pool. */
const allocateXfId = (ss: Stylesheet, style: WriteOnlyStyle): number => {
  // `xfId` is intentionally omitted: leaving it undefined skips the
  // cellStyleXfs bounds check and matches what Excel emits when there
  // is no parent style.
  let xf: CellXf = { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 };
  if (style.font !== undefined) {
    xf = { ...xf, fontId: addFont(ss, style.font), applyFont: true };
  }
  if (style.fill !== undefined) {
    xf = { ...xf, fillId: addFill(ss, style.fill), applyFill: true };
  }
  if (style.border !== undefined) {
    xf = { ...xf, borderId: addBorder(ss, style.border), applyBorder: true };
  }
  if (style.numberFormat !== undefined) {
    xf = { ...xf, numFmtId: addNumFmt(ss, style.numberFormat), applyNumberFormat: true };
  }
  if (style.alignment !== undefined) {
    xf = { ...xf, alignment: style.alignment, applyAlignment: true };
  }
  if (style.protection !== undefined) {
    xf = { ...xf, protection: style.protection, applyProtection: true };
  }
  return addCellXf(ss, xf);
};

interface WorkbookState {
  styles: Stylesheet;
  sst: ReturnType<typeof makeSharedStrings>;
  /** Sheet emit metadata, in addWorksheet order. */
  sheets: Array<{
    title: string;
    sheetId: number;
    bytes: Uint8Array;
  }>;
  /** True once finalize() has been called (further mutations throw). */
  finalised: boolean;
  /** True while a worksheet is open (the next addWorksheet must wait). */
  hasOpenWorksheet: boolean;
}

class WriteOnlyWorksheetImpl implements WriteOnlyWorksheet {
  private readonly state: WorkbookState;
  private readonly sheet: Worksheet;
  /** Next free row index (1-based). appendRow pushes to it then increments. */
  private nextRow = 1;
  private closed = false;
  public readonly title: string;
  /** Pending column widths to set before serialisation. */
  private columnWidths = new Map<number, number>();

  constructor(state: WorkbookState, title: string) {
    this.state = state;
    this.title = title;
    this.sheet = makeWorksheet(title);
  }

  async appendRow(row: WriteOnlyRowItem[]): Promise<void> {
    if (this.closed) throw new OpenXmlIoError('appendRow: worksheet already closed');
    const r = this.nextRow++;
    for (let i = 0; i < row.length; i++) {
      const item = row[i];
      if (item === undefined || item === null) continue;
      const col = i + 1;
      let value: CellValue;
      let style: WriteOnlyStyle | undefined;
      if (item !== null && typeof item === 'object' && 'value' in (item as object)) {
        const wrapped = item as { value: CellValue; style?: WriteOnlyStyle };
        value = wrapped.value;
        style = wrapped.style;
      } else {
        value = item as CellValue;
      }
      const styleId = style ? allocateXfId(this.state.styles, style) : 0;
      setCell(this.sheet, r, col, value, styleId);
    }
  }

  setColumnWidth(col: number, width: number): void {
    if (this.closed) throw new OpenXmlIoError('setColumnWidth: worksheet already closed');
    this.columnWidths.set(col, width);
  }

  async close(): Promise<void> {
    if (this.closed) return;
    this.closed = true;
    // Apply pending column widths.
    for (const [col, width] of this.columnWidths) {
      this.sheet.columnDimensions.set(col, { min: col, max: col, width, customWidth: true });
    }
    // Serialise via the regular worksheet writer. SharedStrings dedup is
    // performed in-place by the writer.
    const bytes = worksheetToBytes(this.sheet, {
      sharedStrings: this.state.sst,
      rels: makeRelationships(),
      registerTable: () => {
        throw new OpenXmlIoError('write-only: table support is not yet wired');
      },
      registerComments: () => {
        throw new OpenXmlIoError('write-only: comments support is not yet wired');
      },
      registerDrawing: () => {
        throw new OpenXmlIoError('write-only: drawings support is not yet wired');
      },
    });
    const sheetId = this.state.sheets.length + 1;
    this.state.sheets.push({ title: this.title, sheetId, bytes });
    this.state.hasOpenWorksheet = false;
  }
}

class WriteOnlyWorkbookImpl implements WriteOnlyWorkbook {
  private readonly sink: XlsxSink;
  private readonly state: WorkbookState;

  constructor(sink: XlsxSink) {
    this.sink = sink;
    const styles = makeStylesheet();
    // Reserve cellXfs[0] for the default (no apply* flags). Unstyled
    // cells point at this slot via styleId=0; user-styled cells start at
    // index 1 so the writer emits an `s="N"` attribute for them.
    addCellXf(styles, defaultCellXf());
    this.state = {
      styles,
      sst: makeSharedStrings(),
      sheets: [],
      finalised: false,
      hasOpenWorksheet: false,
    };
  }

  async addWorksheet(title: string): Promise<WriteOnlyWorksheet> {
    if (this.state.finalised) {
      throw new OpenXmlIoError('addWorksheet: workbook already finalised');
    }
    if (this.state.hasOpenWorksheet) {
      throw new OpenXmlIoError(
        'addWorksheet: previous worksheet still open — call close() before opening the next one',
      );
    }
    const taken = new Set(this.state.sheets.map((s) => s.title));
    validateTitle(title, taken);
    this.state.hasOpenWorksheet = true;
    return new WriteOnlyWorksheetImpl(this.state, title);
  }

  async finalize(): Promise<void> {
    if (this.state.finalised) {
      throw new OpenXmlIoError('finalize: already finalised');
    }
    if (this.state.hasOpenWorksheet) {
      throw new OpenXmlIoError(
        'finalize: a worksheet is still open — call close() before finalising',
      );
    }
    this.state.finalised = true;
    const writer = createZipWriter(this.sink);

    // 1. Worksheets.
    for (const s of this.state.sheets) {
      await writer.addEntry(`xl/worksheets/sheet${s.sheetId}.xml`, s.bytes);
    }

    // 2. Stylesheet.
    await writer.addEntry(ARC_STYLE, stylesheetToBytes(this.state.styles));

    // 3. SharedStrings (only when non-empty).
    if (this.state.sst.entries.length > 0) {
      await writer.addEntry(ARC_SHARED_STRINGS, sharedStringsToBytes(this.state.sst));
    }

    // 4. workbook.xml.
    const workbookXml = serializeWorkbookXml(this.state.sheets);
    await writer.addEntry(ARC_WORKBOOK, new TextEncoder().encode(workbookXml));

    // 5. workbook.xml.rels.
    const wbRels = makeRelationships();
    this.state.sheets.forEach((s, i) => {
      wbRels.rels.push({
        id: `rId${i + 1}`,
        type: `${REL_NS}/worksheet`,
        target: `worksheets/sheet${s.sheetId}.xml`,
      });
    });
    if (this.state.sst.entries.length > 0) {
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
    await writer.addEntry(ARC_WORKBOOK_RELS, relsToBytes(wbRels));

    // 6. root rels.
    const rootRels = makeRelationships();
    rootRels.rels.push({
      id: 'rId1',
      type: `${REL_NS}/officeDocument`,
      target: 'xl/workbook.xml',
    });
    await writer.addEntry(ARC_ROOT_RELS, relsToBytes(rootRels));
    void PKG_REL_NS; // imported for future docProps support

    // 7. [Content_Types].xml.
    const manifest = makeManifest();
    addOverride(manifest, `/${ARC_WORKBOOK}`, XLSX_TYPE);
    for (const s of this.state.sheets) {
      addOverride(manifest, `/xl/worksheets/sheet${s.sheetId}.xml`, WORKSHEET_TYPE);
    }
    addOverride(manifest, `/${ARC_STYLE}`, STYLES_TYPE);
    if (this.state.sst.entries.length > 0) {
      addOverride(manifest, `/${ARC_SHARED_STRINGS}`, SHARED_STRINGS_TYPE);
    }
    await writer.addEntry(ARC_CONTENT_TYPES, manifestToBytes(manifest));

    await writer.finalize();
  }
}

const serializeWorkbookXml = (
  sheets: ReadonlyArray<{ title: string; sheetId: number }>,
): string => {
  const parts: string[] = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<workbook xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}">`,
    '<sheets>',
  ];
  sheets.forEach((s, i) => {
    parts.push(
      `<sheet name="${escapeAttr(s.title)}" sheetId="${s.sheetId}" r:id="rId${i + 1}"/>`,
    );
  });
  parts.push('</sheets></workbook>');
  return parts.join('');
};

/** Open a workbook for streaming write-only output. */
export async function createWriteOnlyWorkbook(
  sink: XlsxSink,
  _opts: WriteOnlyOptions = {},
): Promise<WriteOnlyWorkbook> {
  // The buffered ZIP writer is created lazily inside finalize() so the
  // caller can compose multiple sheets without an early commit. The
  // current implementation accumulates worksheet bytes in memory; the
  // streaming-deflate rewrite (per docs/plan/06-streaming.md §3.4
  // acceptance — 100M cells in 1GB heap) lands as a follow-up.
  return new WriteOnlyWorkbookImpl(sink);
}

