// XML mapping for Font.

import { defineSchema, type Schema } from '../schema/core';
import { SHEET_MAIN_NS } from '../xml/namespaces';
import { ColorSchema } from './colors.schema';
import type { Font } from './fonts';
import { FONT_SCHEMES, UNDERLINE_STYLES, VERT_ALIGNS } from './fonts';

const UNDERLINE_VALUES: readonly string[] = UNDERLINE_STYLES;
const VERT_VALUES: readonly string[] = VERT_ALIGNS;
const SCHEME_VALUES: readonly string[] = FONT_SCHEMES;

export const FontSchema: Schema<Font> = defineSchema<Font>({
  tagname: 'font',
  xmlNs: SHEET_MAIN_NS,
  attrs: {},
  // Element order matches openpyxl's __elements__ tuple for byte-level
  // parity with the typical Excel emit order.
  elements: [
    { kind: 'nested', key: 'name', xmlNs: SHEET_MAIN_NS, primitive: 'string', optional: true },
    { kind: 'nested', key: 'charset', xmlNs: SHEET_MAIN_NS, primitive: 'int', optional: true },
    { kind: 'nested', key: 'family', xmlNs: SHEET_MAIN_NS, primitive: 'int', min: 0, max: 14, optional: true },
    { kind: 'empty', key: 'bold', name: 'b', xmlNs: SHEET_MAIN_NS },
    { kind: 'empty', key: 'italic', name: 'i', xmlNs: SHEET_MAIN_NS },
    { kind: 'empty', key: 'strike', xmlNs: SHEET_MAIN_NS },
    { kind: 'empty', key: 'outline', xmlNs: SHEET_MAIN_NS },
    { kind: 'empty', key: 'shadow', xmlNs: SHEET_MAIN_NS },
    { kind: 'empty', key: 'condense', xmlNs: SHEET_MAIN_NS },
    { kind: 'empty', key: 'extend', xmlNs: SHEET_MAIN_NS },
    { kind: 'object', key: 'color', schema: () => ColorSchema, optional: true },
    { kind: 'nested', key: 'size', name: 'sz', xmlNs: SHEET_MAIN_NS, primitive: 'float', optional: true },
    {
      // ECMA-376 §18.4.13: bare `<u/>` (no @val) means underline=single.
      // Real Excel emits this for hyperlink fonts.
      kind: 'nested',
      key: 'underline',
      name: 'u',
      xmlNs: SHEET_MAIN_NS,
      primitive: 'enum',
      values: UNDERLINE_VALUES,
      optional: true,
      default: 'single',
    },
    {
      kind: 'nested',
      key: 'vertAlign',
      xmlNs: SHEET_MAIN_NS,
      primitive: 'enum',
      values: VERT_VALUES,
      optional: true,
    },
    {
      kind: 'nested',
      key: 'scheme',
      xmlNs: SHEET_MAIN_NS,
      primitive: 'enum',
      values: SCHEME_VALUES,
      optional: true,
    },
  ],
});
