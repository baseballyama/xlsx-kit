// XML mapping for NumberFormat (the per-style <numFmt> element under
// the stylesheet's <numFmts> sequence).

import { defineSchema, type Schema } from '../schema/core';
import { SHEET_MAIN_NS } from '../xml/namespaces';
import type { NumberFormat } from './numbers';

export const NumberFormatSchema: Schema<NumberFormat> = defineSchema<NumberFormat>({
  tagname: 'numFmt',
  xmlNs: SHEET_MAIN_NS,
  attrs: {
    numFmtId: { kind: 'int', min: 0 },
    formatCode: { kind: 'string' },
  },
  elements: [],
});
