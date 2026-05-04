// `docProps/custom.xml` — user-defined document properties. Each
// <property> carries fmtid/pid/name attributes plus a single typed
// value child from the vt: namespace (vt:lpwstr, vt:i4, vt:bool, …).
//
// Per docs/plan/03-foundations.md §6.3. The Schema layer drives the
// <property> element attribute set; the typed-value child is stored as
// a raw XmlNode and the `make*Value` / `read*Value` helpers below cover
// the most common conversions.

import { OpenXmlSchemaError } from '../utils/exceptions';
import { CPROPS_FMTID, CUSTPROPS_NS, parseQName, VTYPES_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { serializeXml } from '../xml/serializer';
import { el, type XmlNode } from '../xml/tree';

export interface CustomProperty {
  /** User-visible property name (must be unique within the workbook). */
  name: string;
  /** OOXML "Property Identifier"; ≥ 2. Auto-allocated if absent on append. */
  pid: number;
  /** Format ID GUID. Defaults to the well-known {D5CDD505-…} on append. */
  fmtid?: string;
  /** Typed value as a raw vt: element. Use the `make*Value` helpers. */
  value: XmlNode;
}

export interface CustomProperties {
  properties: CustomProperty[];
}

export function makeCustomProperties(): CustomProperties {
  return { properties: [] };
}

const PROPERTY_NAME = `{${CUSTPROPS_NS}}property`;
const PROPERTIES_NAME = `{${CUSTPROPS_NS}}Properties`;

// ---- typed-value constructors ----------------------------------------------

const vt = (local: string, text: string): XmlNode => el(`{${VTYPES_NS}}${local}`, {}, [], text);

export function makeStringValue(s: string): XmlNode {
  return vt('lpwstr', s);
}
export function makeAsciiStringValue(s: string): XmlNode {
  return vt('lpstr', s);
}
export function makeIntValue(n: number): XmlNode {
  if (!Number.isInteger(n)) throw new OpenXmlSchemaError(`makeIntValue: ${n} is not an integer`);
  return vt('i4', String(n));
}
export function makeDoubleValue(n: number): XmlNode {
  if (!Number.isFinite(n)) throw new OpenXmlSchemaError(`makeDoubleValue: ${n} is not finite`);
  return vt('r8', String(n));
}
export function makeBoolValue(b: boolean): XmlNode {
  return vt('bool', b ? '1' : '0');
}
export function makeFiletimeValue(iso: string): XmlNode {
  return vt('filetime', iso);
}
export function makeDateValue(iso: string): XmlNode {
  return vt('date', iso);
}

// ---- typed-value readers ---------------------------------------------------

const localNameOf = (n: XmlNode): string => parseQName(n.name).local;

export function readStringValue(v: XmlNode): string | undefined {
  const ln = localNameOf(v);
  if (ln === 'lpwstr' || ln === 'lpstr' || ln === 'bstr') return v.text ?? '';
  return undefined;
}
export function readIntValue(v: XmlNode): number | undefined {
  const ln = localNameOf(v);
  if (
    ln === 'i4' ||
    ln === 'i2' ||
    ln === 'i1' ||
    ln === 'int' ||
    ln === 'uint' ||
    ln === 'ui4' ||
    ln === 'ui2' ||
    ln === 'ui1'
  ) {
    const n = Number.parseInt(v.text ?? '', 10);
    return Number.isFinite(n) ? n : undefined;
  }
  return undefined;
}
export function readDoubleValue(v: XmlNode): number | undefined {
  const ln = localNameOf(v);
  if (ln === 'r4' || ln === 'r8' || ln === 'decimal' || ln === 'cy') {
    const n = Number.parseFloat(v.text ?? '');
    return Number.isFinite(n) ? n : undefined;
  }
  return undefined;
}
export function readBoolValue(v: XmlNode): boolean | undefined {
  if (localNameOf(v) !== 'bool') return undefined;
  const t = (v.text ?? '').toLowerCase();
  if (t === '1' || t === 'true' || t === 't') return true;
  if (t === '0' || t === 'false' || t === 'f') return false;
  return undefined;
}
export function readFiletimeValue(v: XmlNode): string | undefined {
  if (localNameOf(v) === 'filetime') return v.text ?? '';
  return undefined;
}

// ---- collection ops --------------------------------------------------------

const allocatePid = (props: CustomProperties): number => {
  // pid 0 / 1 are reserved by the OPC spec; user pids start at 2.
  const used = new Set<number>();
  for (const p of props.properties) used.add(p.pid);
  let n = 2;
  while (used.has(n)) n++;
  return n;
};

export function appendCustomProperty(
  props: CustomProperties,
  name: string,
  value: XmlNode,
  opts?: { pid?: number; fmtid?: string },
): CustomProperty {
  const pid = opts?.pid ?? allocatePid(props);
  const out: CustomProperty = { name, pid, value };
  if (opts?.fmtid !== undefined) out.fmtid = opts.fmtid;
  props.properties.push(out);
  return out;
}

export function findCustomPropertyByName(props: CustomProperties, name: string): CustomProperty | undefined {
  for (const p of props.properties) if (p.name === name) return p;
  return undefined;
}

// ---- bytes round-trip ------------------------------------------------------

export function customPropsToBytes(p: CustomProperties): Uint8Array {
  const root = el(PROPERTIES_NAME);
  for (const prop of p.properties) {
    const propEl = el(PROPERTY_NAME, {
      fmtid: prop.fmtid ?? CPROPS_FMTID,
      pid: String(prop.pid),
      name: prop.name,
    });
    propEl.children.push(prop.value);
    root.children.push(propEl);
  }
  return serializeXml(root);
}

export function customPropsFromBytes(bytes: Uint8Array | string): CustomProperties {
  const root = parseXml(bytes);
  if (root.name !== PROPERTIES_NAME) {
    throw new OpenXmlSchemaError(`customPropsFromBytes: expected <Properties>, got "${root.name}"`);
  }
  const properties: CustomProperty[] = [];
  for (const propEl of root.children) {
    if (propEl.name !== PROPERTY_NAME) continue;
    const fmtid = propEl.attrs['fmtid'];
    const pidRaw = propEl.attrs['pid'];
    const name = propEl.attrs['name'];
    if (name === undefined || pidRaw === undefined) {
      throw new OpenXmlSchemaError('custom.xml: <property> requires name and pid attributes');
    }
    const pid = Number.parseInt(pidRaw, 10);
    if (!Number.isFinite(pid)) {
      throw new OpenXmlSchemaError(`custom.xml: <property> pid is not an integer (got "${pidRaw}")`);
    }
    const value = propEl.children[0];
    if (value === undefined) {
      throw new OpenXmlSchemaError(`custom.xml: <property name="${name}"> has no typed-value child`);
    }
    const out: CustomProperty = { name, pid, value };
    if (fmtid !== undefined) out.fmtid = fmtid;
    properties.push(out);
  }
  return { properties };
}
