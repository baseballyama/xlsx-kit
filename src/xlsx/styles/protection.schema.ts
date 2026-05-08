// XML mapping for Protection.

import { defineSchema, type Schema } from '../../schema/core';
import { SHEET_MAIN_NS } from '../../xml/namespaces';
import type { Protection } from './protection';

export const ProtectionSchema: Schema<Protection> = defineSchema<Protection>({
  tagname: 'protection',
  xmlNs: SHEET_MAIN_NS,
  attrs: {
    locked: { kind: 'bool', optional: true },
    hidden: { kind: 'bool', optional: true },
  },
  elements: [],
});
