// Cell alignment value object. Mirrors openpyxl/openpyxl/styles/alignment.py.

import { OpenXmlSchemaError } from '../utils/exceptions';

export type HorizontalAlignment =
  | 'general'
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerContinuous'
  | 'distributed';

export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

export const HORIZONTAL_ALIGNMENTS: ReadonlyArray<HorizontalAlignment> = Object.freeze([
  'general',
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
]);
export const VERTICAL_ALIGNMENTS: ReadonlyArray<VerticalAlignment> = Object.freeze([
  'top',
  'center',
  'bottom',
  'justify',
  'distributed',
]);

const HORIZONTAL_SET: ReadonlySet<string> = new Set(HORIZONTAL_ALIGNMENTS);
const VERTICAL_SET: ReadonlySet<string> = new Set(VERTICAL_ALIGNMENTS);

export interface Alignment {
  readonly horizontal?: HorizontalAlignment;
  readonly vertical?: VerticalAlignment;
  /** 0..180 (degrees) OR 255 (vertical stacked text). */
  readonly textRotation?: number;
  readonly wrapText?: boolean;
  readonly shrinkToFit?: boolean;
  /** 0..255 indent levels. */
  readonly indent?: number;
  /** -255..255 relative indent. */
  readonly relativeIndent?: number;
  readonly justifyLastLine?: boolean;
  /** 0 = context-dependent, 1 = LTR, 2 = RTL. */
  readonly readingOrder?: number;
}

export function makeAlignment(opts: Partial<Alignment> = {}): Alignment {
  const out: { -readonly [K in keyof Alignment]: Alignment[K] } = {};
  if (opts.horizontal !== undefined) {
    if (!HORIZONTAL_SET.has(opts.horizontal)) {
      throw new OpenXmlSchemaError(
        `Alignment horizontal must be one of [${HORIZONTAL_ALIGNMENTS.join(', ')}]; got "${opts.horizontal}"`,
      );
    }
    out.horizontal = opts.horizontal;
  }
  if (opts.vertical !== undefined) {
    if (!VERTICAL_SET.has(opts.vertical)) {
      throw new OpenXmlSchemaError(
        `Alignment vertical must be one of [${VERTICAL_ALIGNMENTS.join(', ')}]; got "${opts.vertical}"`,
      );
    }
    out.vertical = opts.vertical;
  }
  if (opts.textRotation !== undefined) {
    const r = opts.textRotation;
    if (!Number.isInteger(r) || r < 0 || (r > 180 && r !== 255)) {
      throw new OpenXmlSchemaError(`Alignment textRotation must be 0..180 or 255; got ${r}`);
    }
    out.textRotation = r;
  }
  if (opts.wrapText !== undefined) out.wrapText = opts.wrapText;
  if (opts.shrinkToFit !== undefined) out.shrinkToFit = opts.shrinkToFit;
  if (opts.indent !== undefined) {
    if (!Number.isFinite(opts.indent) || opts.indent < 0 || opts.indent > 255) {
      throw new OpenXmlSchemaError(`Alignment indent must be 0..255; got ${opts.indent}`);
    }
    out.indent = opts.indent;
  }
  if (opts.relativeIndent !== undefined) {
    const r = opts.relativeIndent;
    if (!Number.isFinite(r) || r < -255 || r > 255) {
      throw new OpenXmlSchemaError(`Alignment relativeIndent must be -255..255; got ${r}`);
    }
    out.relativeIndent = r;
  }
  if (opts.justifyLastLine !== undefined) out.justifyLastLine = opts.justifyLastLine;
  if (opts.readingOrder !== undefined) {
    if (!Number.isFinite(opts.readingOrder) || opts.readingOrder < 0) {
      throw new OpenXmlSchemaError(`Alignment readingOrder must be >= 0; got ${opts.readingOrder}`);
    }
    out.readingOrder = opts.readingOrder;
  }
  return Object.freeze(out);
}

export const DEFAULT_ALIGNMENT: Alignment = makeAlignment();
