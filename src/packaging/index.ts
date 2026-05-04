// Phase 1 §6 packaging entrypoint. Manifest + Relationships are the
// two structural files every OOXML zip carries; doc properties (core,
// app, custom) follow in the next bootstrap-style turn.

export type { DefaultEntry, Manifest, OverrideEntry } from './manifest';
export {
  addDefault,
  addOverride,
  findOverride,
  findOverrideByContentType,
  makeManifest,
  manifestFromBytes,
  manifestToBytes,
} from './manifest';
export type { Relationship, Relationships } from './relationships';
export {
  appendRel,
  findAllByType,
  findById,
  findByType,
  makeRelationships,
  relsFromBytes,
  relsToBytes,
} from './relationships';
