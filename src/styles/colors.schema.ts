// XML mapping for Color. Lives in a sibling file so a build path that
// only reads / mutates colour values without serialising them can drop
// the schema entirely (docs/plan/01-architecture.md §5.3).

import { defineSchema, type Schema } from '../schema/core';
import { SHEET_MAIN_NS } from '../xml/namespaces';
import type { Color } from './colors';

export const ColorSchema: Schema<Color> = defineSchema<Color>({
  tagname: 'color',
  xmlNs: SHEET_MAIN_NS,
  attrs: {
    rgb: { kind: 'string', optional: true },
    // ECMA-376 §18.8.27 documents indices 0..63 + 64 (system fg) + 65
    // (system bg), but Excel emits higher indices (81 = light text on
    // dark background among others) — keep the lower bound and stay
    // permissive on the upper end.
    indexed: { kind: 'int', optional: true, min: 0 },
    theme: { kind: 'int', optional: true, min: 0 },
    auto: { kind: 'bool', optional: true },
    tint: { kind: 'float', optional: true, min: -1, max: 1 },
  },
  elements: [],
});
