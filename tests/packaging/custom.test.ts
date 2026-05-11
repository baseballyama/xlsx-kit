import { describe, expect, it } from 'vitest';
import {
  appendCustomProperty,
  customPropsFromBytes,
  customPropsToBytes,
  findCustomPropertyByName,
  makeBoolValue,
  makeCustomProperties,
  makeDoubleValue,
  makeFiletimeValue,
  makeIntValue,
  makeStringValue,
  readBoolValue,
  readDoubleValue,
  readFiletimeValue,
  readIntValue,
  readStringValue,
} from '../../src/packaging/custom';
import { OpenXmlSchemaError } from '../../src/utils/exceptions';
import { CPROPS_FMTID, CUSTPROPS_NS, VTYPES_NS } from '../../src/xml/namespaces';

describe('customProperties — typed-value helpers', () => {
  it('makeStringValue / readStringValue', () => {
    const v = makeStringValue('John Doe');
    expect(v.name).toBe(`{${VTYPES_NS}}lpwstr`);
    expect(v.text).toBe('John Doe');
    expect(readStringValue(v)).toBe('John Doe');
  });

  it('makeIntValue / readIntValue', () => {
    const v = makeIntValue(42);
    expect(v.name).toBe(`{${VTYPES_NS}}i4`);
    expect(v.text).toBe('42');
    expect(readIntValue(v)).toBe(42);
  });

  it('makeIntValue rejects non-integers', () => {
    expect(() => makeIntValue(1.5)).toThrowError(OpenXmlSchemaError);
  });

  it('makeBoolValue emits 1/0; readBoolValue accepts both shapes', () => {
    expect(makeBoolValue(true).text).toBe('1');
    expect(makeBoolValue(false).text).toBe('0');
    expect(readBoolValue(makeBoolValue(true))).toBe(true);
    expect(readBoolValue(makeBoolValue(false))).toBe(false);
    expect(readBoolValue({ name: `{${VTYPES_NS}}bool`, attrs: {}, children: [], text: 'true' })).toBe(true);
  });

  it('makeDoubleValue / readDoubleValue handle floats', () => {
    const v = makeDoubleValue(3.14);
    expect(v.name).toBe(`{${VTYPES_NS}}r8`);
    expect(readDoubleValue(v)).toBe(3.14);
  });

  it('readers return undefined for non-matching value tags', () => {
    expect(readIntValue(makeStringValue('x'))).toBeUndefined();
    expect(readBoolValue(makeIntValue(1))).toBeUndefined();
    expect(readStringValue(makeIntValue(1))).toBeUndefined();
  });

  it('makeFiletimeValue / readFiletimeValue', () => {
    const iso = '2026-05-04T12:00:00Z';
    expect(readFiletimeValue(makeFiletimeValue(iso))).toBe(iso);
  });
});

describe('customProperties — collection ops', () => {
  it('appendCustomProperty auto-allocates pid starting at 2', () => {
    const p = makeCustomProperties();
    const a = appendCustomProperty(p, 'Editor', makeStringValue('Alice'));
    const b = appendCustomProperty(p, 'Department', makeStringValue('Eng'));
    expect(a.pid).toBe(2);
    expect(b.pid).toBe(3);
  });

  it('skips pids manually present', () => {
    const p = makeCustomProperties();
    appendCustomProperty(p, 'A', makeIntValue(1), { pid: 5 });
    const auto = appendCustomProperty(p, 'B', makeIntValue(2));
    expect(auto.pid).toBe(2);
    appendCustomProperty(p, 'C', makeIntValue(3), { pid: 3 });
    const next = appendCustomProperty(p, 'D', makeIntValue(4));
    expect(next.pid).toBe(4);
  });

  it('findCustomPropertyByName', () => {
    const p = makeCustomProperties();
    appendCustomProperty(p, 'Editor', makeStringValue('Alice'));
    expect(findCustomPropertyByName(p, 'Editor')?.name).toBe('Editor');
    expect(findCustomPropertyByName(p, 'Missing')).toBeUndefined();
  });
});

describe('customProperties — XML round-trip', () => {
  it('serialises one property correctly', () => {
    const p = makeCustomProperties();
    appendCustomProperty(p, 'Editor', makeStringValue('Alice'));
    const xml = new TextDecoder().decode(customPropsToBytes(p));
    expect(xml).toContain(`xmlns="${CUSTPROPS_NS}"`);
    expect(xml).toContain('<property');
    expect(xml).toContain(`fmtid="${CPROPS_FMTID}"`);
    expect(xml).toContain('pid="2"');
    expect(xml).toContain('name="Editor"');
    expect(xml).toContain('<vt:lpwstr>Alice</vt:lpwstr>');
  });

  it('round-trips a mixed-type custom.xml', () => {
    const p = makeCustomProperties();
    appendCustomProperty(p, 'Editor', makeStringValue('Alice'));
    appendCustomProperty(p, 'Year', makeIntValue(2026));
    appendCustomProperty(p, 'Reviewed', makeBoolValue(true));
    appendCustomProperty(p, 'Score', makeDoubleValue(98.5));
    appendCustomProperty(p, 'Updated', makeFiletimeValue('2026-05-04T12:00:00Z'));

    const back = customPropsFromBytes(customPropsToBytes(p));
    const [editor, year, reviewed, score, updated] = back.properties;
    if (!editor || !year || !reviewed || !score || !updated) {
      throw new Error('expected 5 properties to round-trip');
    }
    expect(readStringValue(editor.value)).toBe('Alice');
    expect(readIntValue(year.value)).toBe(2026);
    expect(readBoolValue(reviewed.value)).toBe(true);
    expect(readDoubleValue(score.value)).toBe(98.5);
    expect(readFiletimeValue(updated.value)).toBe('2026-05-04T12:00:00Z');
  });

  it('rejects malformed input (missing pid)', () => {
    const xml = `<?xml version="1.0"?><Properties xmlns="${CUSTPROPS_NS}" xmlns:vt="${VTYPES_NS}"><property fmtid="${CPROPS_FMTID}" name="X"><vt:lpwstr>v</vt:lpwstr></property></Properties>`;
    expect(() => customPropsFromBytes(xml)).toThrowError(OpenXmlSchemaError);
  });

  it('rejects properties without a typed-value child', () => {
    const xml = `<?xml version="1.0"?><Properties xmlns="${CUSTPROPS_NS}" xmlns:vt="${VTYPES_NS}"><property fmtid="${CPROPS_FMTID}" pid="2" name="X"></property></Properties>`;
    expect(() => customPropsFromBytes(xml)).toThrowError(OpenXmlSchemaError);
  });
});
