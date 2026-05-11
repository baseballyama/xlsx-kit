import { describe, expect, it } from 'vitest';
import { fromTree, toTree } from '../../src/schema/serialize';
import {
  type Alignment,
  DEFAULT_ALIGNMENT,
  HORIZONTAL_ALIGNMENTS,
  makeAlignment,
  VERTICAL_ALIGNMENTS,
} from '../../src/styles/alignment';
import { AlignmentSchema } from '../../src/styles/alignment.schema';
import { OpenXmlSchemaError } from '../../src/utils/exceptions';
import { parseXml } from '../../src/xml/parser';
import { serializeXml } from '../../src/xml/serializer';

describe('Alignment', () => {
  it('makeAlignment returns a frozen object', () => {
    const a = makeAlignment({ horizontal: 'left' });
    expect(Object.isFrozen(a)).toBe(true);
  });

  it('omits unset fields entirely', () => {
    expect(makeAlignment({})).toEqual({});
  });

  it('accepts every horizontal / vertical alignment', () => {
    for (const h of HORIZONTAL_ALIGNMENTS) expect(makeAlignment({ horizontal: h }).horizontal).toBe(h);
    for (const v of VERTICAL_ALIGNMENTS) expect(makeAlignment({ vertical: v }).vertical).toBe(v);
  });

  it('rejects unknown alignment values', () => {
    // biome-ignore lint/suspicious/noExplicitAny: bad input on purpose
    expect(() => makeAlignment({ horizontal: 'sideways' as any })).toThrowError(OpenXmlSchemaError);
    // biome-ignore lint/suspicious/noExplicitAny: bad input on purpose
    expect(() => makeAlignment({ vertical: 'middle' as any })).toThrowError(OpenXmlSchemaError);
  });

  it('accepts textRotation in 0..180 and the special 255', () => {
    expect(makeAlignment({ textRotation: 0 }).textRotation).toBe(0);
    expect(makeAlignment({ textRotation: 90 }).textRotation).toBe(90);
    expect(makeAlignment({ textRotation: 180 }).textRotation).toBe(180);
    expect(makeAlignment({ textRotation: 255 }).textRotation).toBe(255);
  });

  it('rejects textRotation in (180, 255) and outside [0, 255]', () => {
    expect(() => makeAlignment({ textRotation: 200 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeAlignment({ textRotation: -1 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeAlignment({ textRotation: 256 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeAlignment({ textRotation: 1.5 })).toThrowError(OpenXmlSchemaError);
  });

  it('enforces indent / relativeIndent / readingOrder ranges', () => {
    expect(() => makeAlignment({ indent: -1 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeAlignment({ indent: 256 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeAlignment({ relativeIndent: -300 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeAlignment({ relativeIndent: 300 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeAlignment({ readingOrder: -1 })).toThrowError(OpenXmlSchemaError);
  });

  it('DEFAULT_ALIGNMENT is the empty Alignment', () => {
    expect(DEFAULT_ALIGNMENT).toEqual({});
    expect(Object.isFrozen(DEFAULT_ALIGNMENT)).toBe(true);
  });

  it('round-trips a fully populated Alignment via the schema', () => {
    const a: Alignment = makeAlignment({
      horizontal: 'center',
      vertical: 'top',
      textRotation: 45,
      wrapText: true,
      shrinkToFit: false,
      indent: 2,
      relativeIndent: -1,
      justifyLastLine: true,
      readingOrder: 1,
    });
    const back = fromTree(parseXml(serializeXml(toTree(a, AlignmentSchema))), AlignmentSchema);
    expect(back).toEqual(a);
  });

  it('emits each set field as an attribute on <alignment>', () => {
    const xml = new TextDecoder().decode(
      serializeXml(toTree(makeAlignment({ horizontal: 'left', wrapText: true }), AlignmentSchema)),
    );
    expect(xml).toContain('horizontal="left"');
    expect(xml).toContain('wrapText="1"');
  });
});
