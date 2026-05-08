import { describe, expect, it } from 'vitest';
import { fromTree, toTree } from '../../../src/schema/serialize';
import { DEFAULT_PROTECTION, makeProtection } from '../../../src/xlsx/styles/protection';
import { ProtectionSchema } from '../../../src/xlsx/styles/protection.schema';
import { parseXml } from '../../../src/xml/parser';
import { serializeXml } from '../../../src/xml/serializer';

describe('Protection', () => {
  it('makeProtection freezes the result', () => {
    expect(Object.isFrozen(makeProtection({ locked: true }))).toBe(true);
  });

  it('omits unset fields', () => {
    expect(makeProtection({})).toEqual({});
  });

  it('DEFAULT_PROTECTION matches Excel: { locked: true, hidden: false }', () => {
    expect(DEFAULT_PROTECTION).toEqual({ locked: true, hidden: false });
    expect(Object.isFrozen(DEFAULT_PROTECTION)).toBe(true);
  });

  it('round-trips both fields via the schema', () => {
    const p = makeProtection({ locked: true, hidden: false });
    const back = fromTree(parseXml(serializeXml(toTree(p, ProtectionSchema))), ProtectionSchema);
    expect(back).toEqual(p);
  });

  it('round-trips when only one field is set', () => {
    const p = makeProtection({ hidden: true });
    const back = fromTree(parseXml(serializeXml(toTree(p, ProtectionSchema))), ProtectionSchema);
    expect(back).toEqual(p);
  });

  it('emits 1/0 for booleans (OOXML convention)', () => {
    const xml = new TextDecoder().decode(
      serializeXml(toTree(makeProtection({ locked: true, hidden: false }), ProtectionSchema)),
    );
    expect(xml).toContain('locked="1"');
    expect(xml).toContain('hidden="0"');
  });
});
