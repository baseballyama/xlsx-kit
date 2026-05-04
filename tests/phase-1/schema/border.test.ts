import { describe, expect, it } from 'vitest';
import { defineSchema, type Schema } from '../../../src/schema/core';
import { fromTree, toTree } from '../../../src/schema/serialize';
import { OpenXmlSchemaError } from '../../../src/utils/exceptions';
import { SHEET_MAIN_NS } from '../../../src/xml/namespaces';
import { parseXml } from '../../../src/xml/parser';
import { serializeXml } from '../../../src/xml/serializer';

// Minimal Border + Side test types — Phase 2 styles will define the
// canonical versions; this fixture proves the schema layer works.

interface Side {
  style?: 'thin' | 'medium' | 'thick' | 'double' | 'hair';
  color?: string;
}

interface Border {
  diagonalUp?: boolean;
  diagonalDown?: boolean;
  outline?: boolean;
  left?: Side;
  right?: Side;
  top?: Side;
  bottom?: Side;
}

const SideSchema: Schema<Side> = defineSchema<Side>({
  tagname: 'side',
  xmlNs: SHEET_MAIN_NS,
  attrs: {
    style: { kind: 'enum', values: ['thin', 'medium', 'thick', 'double', 'hair'], optional: true },
    color: { kind: 'string', optional: true },
  },
  elements: [],
});

const BorderSchema: Schema<Border> = defineSchema<Border>({
  tagname: 'border',
  xmlNs: SHEET_MAIN_NS,
  attrs: {
    diagonalUp: { kind: 'bool', optional: true },
    diagonalDown: { kind: 'bool', optional: true },
    outline: { kind: 'bool', optional: true },
  },
  elements: [
    { kind: 'object', key: 'left', schema: () => SideSchema, optional: true },
    { kind: 'object', key: 'right', schema: () => SideSchema, optional: true },
    { kind: 'object', key: 'top', schema: () => SideSchema, optional: true },
    { kind: 'object', key: 'bottom', schema: () => SideSchema, optional: true },
  ],
});

describe('schema: Side', () => {
  it('round-trips a typed side via XML strings', () => {
    const side: Side = { style: 'thin', color: 'FF000000' };
    const xml = serializeXml(toTree(side, SideSchema));
    const parsed = fromTree(parseXml(xml), SideSchema);
    expect(parsed).toEqual(side);
  });

  it('omits optional missing attributes', () => {
    const side: Side = {};
    const xml = new TextDecoder().decode(serializeXml(toTree(side, SideSchema)));
    expect(xml).toContain('<side');
    expect(xml).not.toContain('style=');
    expect(xml).not.toContain('color=');
  });

  it('rejects an out-of-range enum on parse', () => {
    const xml = `<side xmlns="${SHEET_MAIN_NS}" style="zigzag"/>`;
    expect(() => fromTree(parseXml(xml), SideSchema)).toThrowError(OpenXmlSchemaError);
  });
});

describe('schema: Border', () => {
  it('serializes outline + nested sides into expected XML structure', () => {
    const border: Border = {
      outline: true,
      left: { style: 'thin', color: 'FF112233' },
      bottom: { style: 'medium' },
    };
    const tree = toTree(border, BorderSchema);
    const xml = new TextDecoder().decode(serializeXml(tree));
    expect(xml).toContain('<border');
    expect(xml).toContain('outline="1"');
    expect(xml).toContain('<left style="thin" color="FF112233"/>');
    expect(xml).toContain('<bottom style="medium"/>');
  });

  it('round-trips a fully populated border via XML', () => {
    const border: Border = {
      diagonalUp: true,
      diagonalDown: false,
      outline: true,
      left: { style: 'thin', color: 'FFAABBCC' },
      right: { style: 'medium' },
      top: { style: 'thick' },
      bottom: { style: 'double' },
    };
    const xml = serializeXml(toTree(border, BorderSchema));
    const parsed = fromTree(parseXml(xml), BorderSchema);
    expect(parsed).toEqual(border);
  });

  it('booleans serialise as 1/0 (OOXML convention) and parse back loosely', () => {
    const xml = serializeXml(toTree({ outline: true, diagonalUp: false }, BorderSchema));
    expect(new TextDecoder().decode(xml)).toContain('outline="1"');
    expect(new TextDecoder().decode(xml)).toContain('diagonalUp="0"');

    const looseInput = `<border xmlns="${SHEET_MAIN_NS}" outline="true" diagonalDown="false"/>`;
    const parsed = fromTree(parseXml(looseInput), BorderSchema);
    expect(parsed.outline).toBe(true);
    expect(parsed.diagonalDown).toBe(false);
  });

  it('runs preSerialize / postParse hooks', () => {
    let preCalled = 0;
    let postCalled = 0;
    const HookedSide: Schema<Side> = {
      ...SideSchema,
      preSerialize: (v) => {
        preCalled++;
        // Auto-uppercase color codes during serialise.
        return v.color ? { ...v, color: v.color.toUpperCase() } : v;
      },
      postParse: (v) => {
        postCalled++;
        return v;
      },
    };
    const xml = new TextDecoder().decode(serializeXml(toTree({ color: 'ff0000ff' }, HookedSide)));
    expect(xml).toContain('color="FF0000FF"');
    fromTree(parseXml(xml), HookedSide);
    expect(preCalled).toBe(1);
    expect(postCalled).toBe(1);
  });

  it('toTree throws when a required attribute is undefined', () => {
    const RequiredColor = defineSchema<{ color: string }>({
      tagname: 'side',
      xmlNs: SHEET_MAIN_NS,
      attrs: { color: { kind: 'string' } },
      elements: [],
    });
    // biome-ignore lint/suspicious/noExplicitAny: deliberately exercising the missing-required path
    expect(() => toTree({} as any, RequiredColor)).toThrowError(OpenXmlSchemaError);
  });

  it('fromTree applies attribute defaults when absent', () => {
    const Defaulted = defineSchema<{ outline: boolean }>({
      tagname: 'border',
      xmlNs: SHEET_MAIN_NS,
      attrs: { outline: { kind: 'bool', optional: true, default: true } },
      elements: [],
    });
    const out = fromTree(parseXml(`<border xmlns="${SHEET_MAIN_NS}"/>`), Defaulted);
    expect(out.outline).toBe(true);
  });

  it('handles sequences with a count container', () => {
    interface Foo {
      ids: { v: number }[];
    }
    const ItemSchema = defineSchema<{ v: number }>({
      tagname: 'item',
      xmlNs: SHEET_MAIN_NS,
      attrs: { v: { kind: 'int' } },
      elements: [],
    });
    const FooSchema = defineSchema<Foo>({
      tagname: 'foo',
      xmlNs: SHEET_MAIN_NS,
      attrs: {},
      elements: [
        {
          kind: 'sequence',
          key: 'ids',
          itemName: 'item',
          itemNs: SHEET_MAIN_NS,
          itemSchema: () => ItemSchema,
          container: { name: 'ids', xmlNs: SHEET_MAIN_NS, count: true },
        },
      ],
    });
    const xml = new TextDecoder().decode(serializeXml(toTree({ ids: [{ v: 1 }, { v: 2 }, { v: 3 }] }, FooSchema)));
    expect(xml).toContain('<ids count="3">');
    expect(xml).toContain('<item v="1"/>');
    const round = fromTree(parseXml(xml), FooSchema);
    expect(round['ids']).toEqual([{ v: 1 }, { v: 2 }, { v: 3 }]);
  });
});
