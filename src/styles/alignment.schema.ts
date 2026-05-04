// XML mapping for Alignment.

import { defineSchema, type Schema } from '../schema/core';
import { SHEET_MAIN_NS } from '../xml/namespaces';
import type { Alignment } from './alignment';
import { HORIZONTAL_ALIGNMENTS, VERTICAL_ALIGNMENTS } from './alignment';

const HORIZONTAL_VALUES: readonly string[] = HORIZONTAL_ALIGNMENTS;
const VERTICAL_VALUES: readonly string[] = VERTICAL_ALIGNMENTS;

export const AlignmentSchema: Schema<Alignment> = defineSchema<Alignment>({
  tagname: 'alignment',
  xmlNs: SHEET_MAIN_NS,
  attrs: {
    horizontal: { kind: 'enum', values: HORIZONTAL_VALUES, optional: true },
    vertical: { kind: 'enum', values: VERTICAL_VALUES, optional: true },
    // The makeAlignment constructor enforces the 0..180 ∪ {255} gap; the
    // schema's range check stays loose to avoid rejecting writers that
    // emit 255 for vertical-stacked text.
    textRotation: { kind: 'int', optional: true, min: 0, max: 255 },
    wrapText: { kind: 'bool', optional: true },
    shrinkToFit: { kind: 'bool', optional: true },
    indent: { kind: 'float', optional: true, min: 0, max: 255 },
    relativeIndent: { kind: 'float', optional: true, min: -255, max: 255 },
    justifyLastLine: { kind: 'bool', optional: true },
    readingOrder: { kind: 'float', optional: true, min: 0 },
  },
  elements: [],
});
