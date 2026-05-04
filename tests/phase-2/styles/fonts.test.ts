import { describe, expect, it } from 'vitest';
import { fromTree, toTree } from '../../../src/schema/serialize';
import { makeColor } from '../../../src/styles/colors';
import {
  DEFAULT_FONT,
  FONT_SCHEMES,
  type Font,
  makeFont,
  UNDERLINE_STYLES,
  VERT_ALIGNS,
} from '../../../src/styles/fonts';
import { FontSchema } from '../../../src/styles/fonts.schema';
import { OpenXmlSchemaError } from '../../../src/utils/exceptions';
import { parseXml } from '../../../src/xml/parser';
import { serializeXml } from '../../../src/xml/serializer';

describe('Font value object', () => {
  it('makeFont returns a frozen object', () => {
    expect(Object.isFrozen(makeFont({ name: 'Arial' }))).toBe(true);
  });

  it('omits unset fields entirely', () => {
    expect(makeFont({})).toEqual({});
  });

  it('rejects out-of-range family', () => {
    expect(() => makeFont({ family: -1 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeFont({ family: 15 })).toThrowError(OpenXmlSchemaError);
  });

  it('rejects non-positive size', () => {
    expect(() => makeFont({ size: 0 })).toThrowError(OpenXmlSchemaError);
    expect(() => makeFont({ size: -1 })).toThrowError(OpenXmlSchemaError);
  });

  it('accepts every legal underline / vertAlign / scheme value', () => {
    for (const u of UNDERLINE_STYLES) expect(makeFont({ underline: u }).underline).toBe(u);
    for (const v of VERT_ALIGNS) expect(makeFont({ vertAlign: v }).vertAlign).toBe(v);
    for (const s of FONT_SCHEMES) expect(makeFont({ scheme: s }).scheme).toBe(s);
  });

  it('rejects unknown enum values', () => {
    // biome-ignore lint/suspicious/noExplicitAny: bad input on purpose
    expect(() => makeFont({ underline: 'wavy' as any })).toThrowError(OpenXmlSchemaError);
    // biome-ignore lint/suspicious/noExplicitAny: bad input on purpose
    expect(() => makeFont({ vertAlign: 'middle' as any })).toThrowError(OpenXmlSchemaError);
  });

  it('freezes nested color', () => {
    const f = makeFont({ color: { rgb: 'FF000000' } });
    expect(f.color && Object.isFrozen(f.color)).toBe(true);
  });

  it('DEFAULT_FONT matches Excel: Calibri 11 minor scheme theme=1', () => {
    expect(DEFAULT_FONT.name).toBe('Calibri');
    expect(DEFAULT_FONT.size).toBe(11);
    expect(DEFAULT_FONT.family).toBe(2);
    expect(DEFAULT_FONT.scheme).toBe('minor');
    expect(DEFAULT_FONT.color).toEqual({ theme: 1 });
  });
});

describe('FontSchema XML round-trip', () => {
  it('round-trips a minimal Font', () => {
    const f = makeFont({ name: 'Calibri', size: 11 });
    const back = fromTree(parseXml(serializeXml(toTree(f, FontSchema))), FontSchema);
    expect(back).toEqual(f);
  });

  it('round-trips a fully-populated Font', () => {
    const f: Font = makeFont({
      name: 'Calibri',
      charset: 1,
      family: 2,
      size: 11,
      color: makeColor({ rgb: 'FF000000' }),
      bold: true,
      italic: true,
      strike: true,
      outline: true,
      shadow: true,
      condense: true,
      extend: true,
      underline: 'double',
      vertAlign: 'superscript',
      scheme: 'minor',
    });
    const back = fromTree(parseXml(serializeXml(toTree(f, FontSchema))), FontSchema);
    expect(back).toEqual(f);
  });

  it('round-trips DEFAULT_FONT', () => {
    const back = fromTree(parseXml(serializeXml(toTree(DEFAULT_FONT, FontSchema))), FontSchema);
    expect(back).toEqual(DEFAULT_FONT);
  });

  it('emits each set field as a nested element with val attribute', () => {
    const xml = new TextDecoder().decode(
      serializeXml(toTree(makeFont({ name: 'Arial', size: 12, bold: true }), FontSchema)),
    );
    expect(xml).toContain('<name val="Arial"/>');
    expect(xml).toContain('<sz val="12"/>');
    expect(xml).toContain('<b/>');
  });

  it('does not emit empty markers when the value is false or absent', () => {
    // italic is omitted entirely; bold is present-as-false. Both must
    // round-trip to "no marker" output.
    const xml = new TextDecoder().decode(serializeXml(toTree(makeFont({ name: 'Arial', bold: false }), FontSchema)));
    expect(xml).not.toContain('<b/>');
    expect(xml).not.toContain('<i/>');
  });

  it('parses openpyxl-style XML with mixed attrs / empty markers', () => {
    const xml = `<?xml version="1.0"?><font xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><name val="Calibri"/><sz val="11"/><b/><color theme="1"/></font>`;
    const f = fromTree(parseXml(xml), FontSchema);
    expect(f).toEqual({
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { theme: 1 },
    });
  });
});
