// XML mapping for Fill. Pattern and gradient variants have their own
// Schema; the wrapper `<fill>` element hand-rolls the
// discriminated-union read/write since the schema layer doesn't yet
// have a tagname-discriminated union element kind.

import { defineSchema, type Schema } from '../schema/core';
import { fromTree, toTree } from '../schema/serialize';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { qname, SHEET_MAIN_NS } from '../xml/namespaces';
import { el, type XmlNode } from '../xml/tree';
import { ColorSchema } from './colors.schema';
import { type Fill, type GradientFill, type GradientStop, PATTERN_TYPES, type PatternFill } from './fills';

const PATTERN_TYPE_VALUES: readonly string[] = PATTERN_TYPES;

const StopSchema: Schema<GradientStop> = defineSchema<GradientStop>({
  tagname: 'stop',
  xmlNs: SHEET_MAIN_NS,
  attrs: {
    position: { kind: 'float', min: 0, max: 1 },
  },
  elements: [{ kind: 'object', key: 'color', schema: () => ColorSchema }],
});

const PatternFillSchema: Schema<PatternFill> = defineSchema<PatternFill>({
  tagname: 'patternFill',
  xmlNs: SHEET_MAIN_NS,
  // `kind` is internal-only and never emitted; postParse re-adds it on read.
  attrs: {
    patternType: { kind: 'enum', values: PATTERN_TYPE_VALUES, optional: true },
  },
  elements: [
    { kind: 'object', key: 'fgColor', schema: () => ColorSchema, optional: true },
    { kind: 'object', key: 'bgColor', schema: () => ColorSchema, optional: true },
  ],
  postParse: (v) => ({ ...v, kind: 'pattern' }),
});

const GradientFillSchema: Schema<GradientFill> = defineSchema<GradientFill>({
  tagname: 'gradientFill',
  xmlNs: SHEET_MAIN_NS,
  attrs: {
    type: { kind: 'enum', values: ['linear', 'path'], optional: true, default: 'linear' },
    degree: { kind: 'float', optional: true },
    left: { kind: 'float', optional: true },
    right: { kind: 'float', optional: true },
    top: { kind: 'float', optional: true },
    bottom: { kind: 'float', optional: true },
  },
  elements: [
    {
      kind: 'sequence',
      key: 'stops',
      itemName: 'stop',
      itemNs: SHEET_MAIN_NS,
      itemSchema: () => StopSchema,
    },
  ],
  postParse: (v) => ({ ...v, kind: 'gradient' }),
});

// ---- <fill> wrapper round-trip --------------------------------------------

const FILL_TAG = qname(SHEET_MAIN_NS, 'fill');
const PATTERN_FILL_TAG = qname(SHEET_MAIN_NS, 'patternFill');
const GRADIENT_FILL_TAG = qname(SHEET_MAIN_NS, 'gradientFill');

/** Build the `<fill>…</fill>` envelope around either variant. */
export function fillToTree(fill: Fill): XmlNode {
  const inner = fill.kind === 'gradient' ? toTree(fill, GradientFillSchema) : toTree(fill, PatternFillSchema);
  return el(FILL_TAG, {}, [inner]);
}

/** Inverse of {@link fillToTree}. Reads `<fill>` and dispatches on the child tag. */
export function fillFromTree(node: XmlNode): Fill {
  if (node.name !== FILL_TAG) {
    throw new OpenXmlSchemaError(`fillFromTree: expected <fill>, got "${node.name}"`);
  }
  const child = node.children[0];
  if (child === undefined) {
    throw new OpenXmlSchemaError('fillFromTree: <fill> has no child element');
  }
  if (child.name === PATTERN_FILL_TAG) return fromTree(child, PatternFillSchema);
  if (child.name === GRADIENT_FILL_TAG) return fromTree(child, GradientFillSchema);
  throw new OpenXmlSchemaError(`fillFromTree: unknown Fill variant "${child.name}"`);
}
