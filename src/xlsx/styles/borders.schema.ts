// XML mapping for Border + Side. Sibling file so non-serialising
// callers can avoid pulling in the schema layer.

import { defineSchema, type Schema } from '../../schema/core';
import { SHEET_MAIN_NS } from '../../xml/namespaces';
import type { Border, Side } from './borders';
import { SIDE_STYLES } from './borders';
import { ColorSchema } from './colors.schema';

// Schema enums require `readonly string[]`; SideStyle is a string-literal
// subset, so a single widening cast gets us through TS' invariance check
// without leaking any new type machinery into call sites.
const SIDE_STYLE_VALUES: readonly string[] = SIDE_STYLES;

export const SideSchema: Schema<Side> = defineSchema<Side>({
  tagname: 'side',
  xmlNs: SHEET_MAIN_NS,
  attrs: {
    style: { kind: 'enum', values: SIDE_STYLE_VALUES, optional: true },
  },
  elements: [{ kind: 'object', key: 'color', schema: () => ColorSchema, optional: true }],
});

export const BorderSchema: Schema<Border> = defineSchema<Border>({
  tagname: 'border',
  xmlNs: SHEET_MAIN_NS,
  attrs: {
    diagonalUp: { kind: 'bool', optional: true },
    diagonalDown: { kind: 'bool', optional: true },
    outline: { kind: 'bool', optional: true },
  },
  elements: [
    { kind: 'object', key: 'left', schema: () => SideSchema, optional: true },
    { kind: 'object', key: 'right', schema: () => SideSchema, optional: true },
    { kind: 'object', key: 'top', schema: () => SideSchema, optional: true },
    { kind: 'object', key: 'bottom', schema: () => SideSchema, optional: true },
    { kind: 'object', key: 'diagonal', schema: () => SideSchema, optional: true },
    { kind: 'object', key: 'vertical', schema: () => SideSchema, optional: true },
    { kind: 'object', key: 'horizontal', schema: () => SideSchema, optional: true },
  ],
});
