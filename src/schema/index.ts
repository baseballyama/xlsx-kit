// Schema layer entrypoint. The runtime walker is one (de)serialise per
// direction, so the surface stays tiny on purpose.

export type { AttrDef, ElementDef, Primitive, Schema } from './core';
export { defineSchema } from './core';
export { fromTree, toTree } from './serialize';
