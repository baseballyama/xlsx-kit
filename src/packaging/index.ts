// Phase 1 §6 packaging entrypoint. Manifest + Relationships are the
// two structural files every OOXML zip carries; doc properties (core,
// app, custom) follow in the next bootstrap-style turn.

export type { CoreProperties } from './core';
export {
  corePropsFromBytes,
  corePropsToBytes,
  makeCoreProperties,
  setWorkbookCategory,
  setWorkbookCreator,
  setWorkbookDescription,
  setWorkbookKeywords,
  setWorkbookLastModifiedBy,
  setWorkbookSubject,
  setWorkbookTitle,
} from './core';
export type { CustomProperties, CustomProperty } from './custom';
export {
  appendCustomProperty,
  customPropsFromBytes,
  customPropsToBytes,
  findCustomPropertyByName,
  makeAsciiStringValue,
  makeBoolValue,
  makeCustomProperties,
  makeDateValue,
  makeDoubleValue,
  makeFiletimeValue,
  makeIntValue,
  makeStringValue,
  readBoolValue,
  readDoubleValue,
  readFiletimeValue,
  readIntValue,
  readStringValue,
} from './custom';
export type { ExtendedProperties } from './extended';
export {
  extendedPropsFromBytes,
  extendedPropsToBytes,
  makeExtendedProperties,
  setWorkbookAppVersion,
  setWorkbookApplication,
  setWorkbookCompany,
  setWorkbookHyperlinkBase,
  setWorkbookManager,
} from './extended';
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
