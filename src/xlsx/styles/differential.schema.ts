// XML mapping for DifferentialStyle. The <dxf> element contains the
// same children as a regular Font / Fill / Border / Alignment /
// Protection / NumberFormat block, in a fixed order documented by
// openpyxl's __elements__ tuple.

import { defineSchema, type Schema } from '../../schema/core';
import { SHEET_MAIN_NS } from '../../xml/namespaces';
import { AlignmentSchema } from './alignment.schema';
import { BorderSchema } from './borders.schema';
import type { DifferentialStyle } from './differential';
import { FontSchema } from './fonts.schema';
import { NumberFormatSchema } from './numbers.schema';
import { ProtectionSchema } from './protection.schema';

export const DifferentialStyleSchema: Schema<DifferentialStyle> = defineSchema<DifferentialStyle>({
  tagname: 'dxf',
  xmlNs: SHEET_MAIN_NS,
  attrs: {},
  elements: [
    { kind: 'object', key: 'font', schema: () => FontSchema, optional: true },
    { kind: 'object', key: 'numFmt', schema: () => NumberFormatSchema, optional: true },
    // Fill is not directly a Schema target — its <fill> envelope is
    // hand-rolled via fillToTree / fillFromTree. For the DXF round-trip
    // we treat it as a raw subtree if present; conditional formatting
    // typically embeds the wrapper directly.
    { kind: 'raw', key: 'fill', xmlNs: SHEET_MAIN_NS, name: 'fill', optional: true },
    { kind: 'object', key: 'alignment', schema: () => AlignmentSchema, optional: true },
    { kind: 'object', key: 'border', schema: () => BorderSchema, optional: true },
    { kind: 'object', key: 'protection', schema: () => ProtectionSchema, optional: true },
  ],
});
